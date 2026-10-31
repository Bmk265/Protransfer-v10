import io, json, sqlite3, secrets, datetime, os
from pathlib import Path
from contextlib import closing
import bcrypt
from fastapi import FastAPI, Request, Response, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook

DB_PATH = Path(__file__).parent / "protransfer.db"

def init_db():
    with closing(sqlite3.connect(DB_PATH)) as db:
        c=db.cursor()
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY,password_hash TEXT,full_name TEXT,role TEXT,department TEXT,sub_department TEXT,plant TEXT,status TEXT DEFAULT 'pending',created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,username TEXT,expires TEXT);
        CREATE TABLE IF NOT EXISTS departments(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS plants(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS org(id INT PRIMARY KEY DEFAULT 1,name TEXT DEFAULT 'TCL Pvt Ltd - Gooty Plant (1400)');
        CREATE TABLE IF NOT EXISTS tcodes(code TEXT PRIMARY KEY,title TEXT,description TEXT,color TEXT,icon TEXT,allowed_depts TEXT,active INT DEFAULT 1,created_at TEXT,updated_at TEXT,created_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_schemas(code TEXT PRIMARY KEY,fields TEXT,layout TEXT,updated_at TEXT,updated_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_records(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,data TEXT,status TEXT DEFAULT 'active',created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_layouts(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,layout_code TEXT UNIQUE,name TEXT,title TEXT,is_default INT DEFAULT 0,fields TEXT,layout_config TEXT,description TEXT,created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS access_grants(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT,tcode_code TEXT,layout_code TEXT,granted_by TEXT,granted_at TEXT,UNIQUE(username,tcode_code,layout_code));
        """)
        for d in ["Production","Quality Control","Final QC","Commercial","Planning","Store"]:
            c.execute("INSERT OR IGNORE INTO departments VALUES(?,?,?)",(d,1,datetime.datetime.now().isoformat()))
        for p in ["Gooty","Plant 1400","HYD","TCL Corporate"]:
            c.execute("INSERT OR IGNORE INTO plants VALUES(?,?,?)",(p,1,datetime.datetime.now().isoformat()))
        c.execute("INSERT OR IGNORE INTO org(id,name) VALUES(1,'TCL Pvt Ltd - Gooty Plant (1400)')")
        defaults=[
            ("NCR0","Non Conformance Records","Quality issues, deviations","#ef4444","ri-error-warning-line","Quality Control,Final QC,Production"),
            ("GCH0","Green Card Records","Green card issuance","#22c55e","ri-checkbox-circle-line","Quality Control,Production"),
            ("PROD","Production Handover","Outer Sheathing","#3b82f6","ri-send-plane-line","Production,Quality Control,Final QC"),
            ("QCH0","Quality Check Records","QC to Final QC","#8b5cf6","ri-shield-check-line","Quality Control,Final QC"),
            ("STO0","Store Transfer","Material issue","#f59e0b","ri-store-2-line","Store,Commercial,Planning"),
            ("PLN0","Planning Details","Master planning from Excel - Job No master","#06b6d4","ri-calendar-check-line","Planning,Production,Commercial,Quality Control,Final QC,Store"),
        ]
        for code,title,desc,color,icon,depts in defaults:
            c.execute("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",(code,title,desc,color,icon,depts,1,datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat(),"system"))
        ncr=json.dumps([{"key":"date","label":"Date","type":"date","required":True,"width":1},{"key":"job_no","label":"Job No / SAP No","type":"text","required":True,"width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"material","label":"Material Description","type":"textarea","width":2},{"key":"nc_type","label":"NC Type","type":"select","options":["Dimension","Visual","Mechanical","Other"],"width":1},{"key":"description","label":"NC Description","type":"textarea","width":3}])
        gch=json.dumps([{"key":"date","label":"Green Card Date","type":"date","width":1},{"key":"job_no","label":"Job No","type":"text","width":1},{"key":"drum_code","label":"Drum Code","type":"text","width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"qc_status","label":"QC Status","type":"select","options":["Pass","Hold","Reject"],"width":1},{"key":"remarks","label":"Remarks","type":"textarea","width":3}])
        pln=json.dumps([{"key":"job_no","label":"Job No","type":"text","required":True,"width":1},{"key":"item_no","label":"Item No","type":"text","width":1},{"key":"sap_no","label":"SAP No","type":"text","width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"material_description","label":"Material Description","type":"textarea","width":2},{"key":"specified_length","label":"Specified Length","type":"number","width":1},{"key":"drum_size","label":"Drum Size","type":"text","width":1},{"key":"plan_date","label":"Plan Date","type":"date","width":1},{"key":"remarks","label":"Remarks","type":"textarea","width":2}])
        c.execute("INSERT OR IGNORE INTO tcode_schemas VALUES(?,?,?,?,?)",("NCR0",ncr,json.dumps({"columns":3}),datetime.datetime.now().isoformat(),"system"))
        c.execute("INSERT OR IGNORE INTO tcode_schemas VALUES(?,?,?,?,?)",("GCH0",gch,json.dumps({"columns":3}),datetime.datetime.now().isoformat(),"system"))
        c.execute("INSERT OR IGNORE INTO tcode_schemas VALUES(?,?,?,?,?)",("PLN0",pln,json.dumps({"columns":3}),datetime.datetime.now().isoformat(),"system"))
        c.execute("INSERT OR IGNORE INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",('NCR0','NCR0-001','Default','Default NCR View',1,ncr,json.dumps({"columns":3}),"Default", 'system','System',datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        c.execute("INSERT OR IGNORE INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",('GCH0','GCH0-001','Default','Default GCH View',1,gch,json.dumps({"columns":3}),"Default", 'system','System',datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        c.execute("INSERT OR IGNORE INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",('PLN0','PLN0-001','Default','Default Planning View',1,pln,json.dumps({"columns":3}),"Default master", 'system','System',datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        ph=bcrypt.hashpw("Admin@123".encode(),bcrypt.gensalt()).decode()
        c.execute("INSERT OR IGNORE INTO users VALUES(?,?,?,?,?,?,?,?,?,?)",("admin",ph,"System Admin","admin","Planning","","Gooty","approved",datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        db.commit()

init_db()
def get_db():
    db=sqlite3.connect(DB_PATH); db.row_factory=sqlite3.Row; return db
def now(): return datetime.datetime.now().isoformat()
def get_user(req:Request):
    t=req.cookies.get("pt")
    if not t: return None
    with closing(get_db()) as d:
        s=d.execute("SELECT * FROM sessions WHERE token=?",(t,)).fetchone()
        if not s: return None
        u=d.execute("SELECT * FROM users WHERE username=?",(s["username"],)).fetchone()
        if not u or u["status"]!="approved": return None
        return dict(u)
def need(req:Request):
    u=get_user(req)
    if not u: raise HTTPException(401,"login")
    return u
def need_admin(req:Request):
    u=need(req)
    if u["role"]!="admin": raise HTTPException(403,"admin")
    return u

app=FastAPI()
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_methods=["*"],allow_headers=["*"])

@app.post("/api/login")
def login(res:Response,username:str=Form(...),password:str=Form(...)):
    username=username.lower().strip()
    with closing(get_db()) as d:
        u=d.execute("SELECT * FROM users WHERE username=?",(username,)).fetchone()
        if not u or not bcrypt.checkpw(password.encode(),u["password_hash"].encode()): raise HTTPException(401,"invalid")
        if u["status"]!="approved": raise HTTPException(403,"not approved")
        tok=secrets.token_urlsafe(32); exp=(datetime.datetime.now()+datetime.timedelta(days=30)).isoformat()
        d.execute("INSERT INTO sessions VALUES(?,?,?)",(tok,username,exp)); d.commit()
    res.set_cookie("pt",tok,httponly=True,samesite="lax",max_age=2592000)
    return {"ok":True}

@app.post("/api/logout")
def logout(res:Response,req:Request):
    with closing(get_db()) as d: d.execute("DELETE FROM sessions WHERE token=?",(req.cookies.get("pt"),)); d.commit()
    res.delete_cookie("pt"); return {"ok":True}

@app.get("/api/me")
def me(req:Request): return need(req)
@app.get("/api/org")
def get_org(req:Request): need(req); return get_db().execute("SELECT * FROM org WHERE id=1").fetchone()
@app.post("/api/org")
def set_org(req:Request,name:str=Form(...)): need_admin(req); conn=get_db(); conn.execute("UPDATE org SET name=? WHERE id=1",(name,)); conn.commit(); return {"ok":True}

# ADMIN USERS
@app.get("/api/admin/users")
def admin_list_users(req:Request):
    need_admin(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT * FROM users ORDER BY created_at DESC").fetchall()]
        return rows

@app.post("/api/admin/create-user")
def admin_create_user(req:Request,username:str=Form(...),password:str=Form(...),full_name:str=Form(...),department:str=Form(...),plant:str=Form(...),role:str=Form(...)):
    need_admin(req); username=username.lower().strip()
    with closing(get_db()) as d:
        if d.execute("SELECT 1 FROM users WHERE username=?",(username,)).fetchone(): raise HTTPException(400,f"User {username} exists")
        ph=bcrypt.hashpw(password.encode(),bcrypt.gensalt()).decode()
        d.execute("INSERT INTO users VALUES(?,?,?,?,?,?,?,?,?,?)",(username,ph,full_name,role,department,"",plant,"approved",now(),now())); d.commit()
    return {"ok":True}

@app.post("/api/admin/update-user")
def admin_update_user(req:Request,username:str=Form(...),full_name:str=Form(...),department:str=Form(...),plant:str=Form(...),role:str=Form(...),status:str=Form(...)):
    need_admin(req)
    with closing(get_db()) as d:
        d.execute("UPDATE users SET full_name=?,department=?,plant=?,role=?,status=?,updated_at=? WHERE username=?",(full_name,department,plant,role,status,now(),username.lower().strip())); d.commit()
    return {"ok":True}

@app.post("/api/admin/delete-user")
def admin_delete_user(req:Request,username:str=Form(...)):
    need_admin(req); username=username.lower().strip()
    if username==need(req)["username"]: raise HTTPException(400,"Cannot delete yourself")
    with closing(get_db()) as d:
        d.execute("DELETE FROM users WHERE username=?",(username,)); d.execute("DELETE FROM sessions WHERE username=?",(username,)); d.execute("DELETE FROM access_grants WHERE username=?",(username,)); d.commit()
    return {"ok":True}

@app.post("/api/admin/reset-password")
def admin_reset_password(req:Request,username:str=Form(...),new_password:str=Form(...)):
    need_admin(req); ph=bcrypt.hashpw(new_password.encode(),bcrypt.gensalt()).decode()
    with closing(get_db()) as d: d.execute("UPDATE users SET password_hash=? WHERE username=?",(ph,username.lower().strip())); d.commit()
    return {"ok":True}

# DEPARTMENTS
@app.get("/api/departments")
def depts(req:Request): need(req); return [dict(r) for r in get_db().execute("SELECT * FROM departments ORDER BY name").fetchall()]
@app.get("/api/admin/departments")
def admin_depts(req:Request): need_admin(req); return [dict(r) for r in get_db().execute("SELECT * FROM departments ORDER BY name").fetchall()]
@app.post("/api/admin/departments/create")
def create_dept(req:Request,name:str=Form(...)):
    need_admin(req); name=name.strip()
    with closing(get_db()) as d: d.execute("INSERT OR REPLACE INTO departments VALUES(?,?,?)",(name,1,now())); d.commit()
    return {"ok":True}
@app.post("/api/admin/departments/delete")
def delete_dept(req:Request,name:str=Form(...)):
    need_admin(req)
    with closing(get_db()) as d: d.execute("DELETE FROM departments WHERE name=?",(name,)); d.commit()
    return {"ok":True}
@app.post("/api/admin/departments/toggle")
def toggle_dept(req:Request,name:str=Form(...)):
    need_admin(req)
    with closing(get_db()) as d:
        r=d.execute("SELECT active FROM departments WHERE name=?",(name,)).fetchone()
        if r: d.execute("UPDATE departments SET active=? WHERE name=?",(0 if r["active"]==1 else 1,name)); d.commit()
    return {"ok":True}

# PLANTS
@app.get("/api/plants")
def plants(req:Request): need(req); return [dict(r) for r in get_db().execute("SELECT * FROM plants ORDER BY name").fetchall()]
@app.get("/api/admin/plants")
def admin_plants(req:Request): need_admin(req); return [dict(r) for r in get_db().execute("SELECT * FROM plants ORDER BY name").fetchall()]
@app.post("/api/admin/plants/create")
def create_plant(req:Request,name:str=Form(...)):
    need_admin(req)
    with closing(get_db()) as d: d.execute("INSERT OR REPLACE INTO plants VALUES(?,?,?)",(name.strip(),1,now())); d.commit()
    return {"ok":True}
@app.post("/api/admin/plants/delete")
def delete_plant(req:Request,name:str=Form(...)):
    need_admin(req)
    with closing(get_db()) as d: d.execute("DELETE FROM plants WHERE name=?",(name,)); d.commit()
    return {"ok":True}
@app.post("/api/admin/plants/toggle")
def toggle_plant(req:Request,name:str=Form(...)):
    need_admin(req)
    with closing(get_db()) as d:
        r=d.execute("SELECT active FROM plants WHERE name=?",(name,)).fetchone()
        if r: d.execute("UPDATE plants SET active=? WHERE name=?",(0 if r["active"]==1 else 1,name)); d.commit()
    return {"ok":True}

# TCODES ADMIN
@app.get("/api/tcodes")
def list_tcodes(req:Request):
    u=need(req)
    with closing(get_db()) as d:
        all_c=[dict(r) for r in d.execute("SELECT * FROM tcodes WHERE active=1 ORDER BY code").fetchall()]
        for tc in all_c: tc["layout_count"]=d.execute("SELECT COUNT(*) c FROM tcode_layouts WHERE tcode_code=?",(tc["code"],)).fetchone()["c"]
        if u["role"]=="admin": return all_c
        grants=[dict(r) for r in d.execute("SELECT tcode_code FROM access_grants WHERE username=?",(u["username"],)).fetchall()]
        granted=set(g["tcode_code"] for g in grants)
        return [tc for tc in all_c if u["department"] in [x.strip() for x in tc["allowed_depts"].split(",")] or tc["code"] in granted]

@app.get("/api/tcodes/all")
def list_tcodes_all(req:Request):
    need_admin(req)
    with closing(get_db()) as d:
        all_c=[dict(r) for r in d.execute("SELECT * FROM tcodes ORDER BY code").fetchall()]
        for tc in all_c: tc["layout_count"]=d.execute("SELECT COUNT(*) c FROM tcode_layouts WHERE tcode_code=?",(tc["code"],)).fetchone()["c"]
        return all_c

@app.post("/api/admin/tcodes/create")
def create_tcode(req:Request,code:str=Form(...),title:str=Form(...),description:str=Form(...),color:str=Form(...),icon:str=Form(...),allowed_depts:str=Form(...)):
    need_admin(req); code=code.upper().strip()
    with closing(get_db()) as d:
        if d.execute("SELECT 1 FROM tcodes WHERE code=?",(code,)).fetchone(): raise HTTPException(400,"T-Code exists")
        d.execute("INSERT INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",(code,title,description,color or "#3b82f6",icon or "ri-apps-line",allowed_depts,1,now(),now(),need(req)["username"]))
        # create default schema + layout
        default_fields=json.dumps([{"key":"job_no","label":"Job No / SAP No","type":"text","required":True,"width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"remarks","label":"Remarks","type":"textarea","width":3}])
        d.execute("INSERT INTO tcode_schemas VALUES(?,?,?,?,?)",(code,default_fields,json.dumps({"columns":3}),now(),need(req)["username"]))
        d.execute("INSERT INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(code,f"{code}-001","Default",f"Default {title} View",1,default_fields,json.dumps({"columns":3}),f"Default for {code}",need(req)["username"],need(req)["full_name"],now(),now()))
        d.commit()
    return {"ok":True, "code":code}

@app.post("/api/admin/tcodes/update")
def update_tcode(req:Request,code:str=Form(...),title:str=Form(...),description:str=Form(...),color:str=Form(...),icon:str=Form(...),allowed_depts:str=Form(...),active:int=Form(...)):
    need_admin(req)
    with closing(get_db()) as d: d.execute("UPDATE tcodes SET title=?,description=?,color=?,icon=?,allowed_depts=?,active=?,updated_at=? WHERE code=?",(title,description,color,icon,allowed_depts,active,now(),code.upper().strip())); d.commit()
    return {"ok":True}

@app.post("/api/admin/tcodes/delete")
def delete_tcode(req:Request,code:str=Form(...)):
    need_admin(req); code=code.upper().strip()
    if code in ["PLN0","NCR0","GCH0"]: raise HTTPException(400,"Cannot delete system T-Codes")
    with closing(get_db()) as d: d.execute("DELETE FROM tcodes WHERE code=?",(code,)); d.execute("DELETE FROM tcode_schemas WHERE code=?",(code,)); d.execute("DELETE FROM tcode_layouts WHERE tcode_code=?",(code,)); d.execute("DELETE FROM tcode_records WHERE tcode_code=?",(code,)); d.commit()
    return {"ok":True}

@app.get("/api/tcode/{code}/layouts")
def list_layouts(code:str,req:Request):
    u=need(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT * FROM tcode_layouts WHERE tcode_code=? ORDER BY is_default DESC, layout_code ASC",(code,)).fetchall()]
        for r in rows: 
            try: r["fields"]=json.loads(r["fields"])
            except: r["fields"]=[]
            r["layout_config"]=json.loads(r["layout_config"]) if r["layout_config"] else {}
        if u["role"]=="admin": return rows
        grants=[dict(r) for r in d.execute("SELECT layout_code FROM access_grants WHERE username=? AND tcode_code=?",(u["username"],code)).fetchall()]
        if not grants: return rows
        if any(g["layout_code"]=="" for g in grants): return rows
        allowed=set(g["layout_code"] for g in grants if g["layout_code"])
        return [r for r in rows if r["layout_code"] in allowed]

@app.post("/api/tcode/{code}/layouts")
async def create_layout(code:str,req:Request):
    u=need(req); b=await req.json()
    layout_code=b.get("layout_code","").upper().strip()
    if not layout_code:
        with closing(get_db()) as d:
            cnt=d.execute("SELECT COUNT(*) c FROM tcode_layouts WHERE tcode_code=?",(code,)).fetchone()["c"]
            layout_code=f"{code}-{cnt+1:03d}"
    with closing(get_db()) as d:
        try:
            d.execute("INSERT INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(code,layout_code,b.get("name",""),b.get("title",""),0,json.dumps(b["fields"]),json.dumps(b.get("layout_config",{"columns":3})),b.get("description",""),u["username"],u["full_name"],now(),now()))
            d.commit()
        except sqlite3.IntegrityError:
            raise HTTPException(400,f"Layout Code {layout_code} exists")
        return {"ok":True,"layout_code":layout_code}

@app.delete("/api/tcode/{code}/layouts/{lid}")
def del_layout(code:str,lid:int,req:Request):
    u=need(req)
    with closing(get_db()) as d:
        r=d.execute("SELECT * FROM tcode_layouts WHERE id=?",(lid,)).fetchone()
        if not r: raise HTTPException(404,"not found")
        if r["is_default"]==1: raise HTTPException(400,"Cannot delete default")
        if r["created_by"]!=u["username"] and u["role"]!="admin": raise HTTPException(403,"not owner")
        d.execute("DELETE FROM tcode_layouts WHERE id=?",(lid,)); d.commit()
    return {"ok":True}

@app.post("/api/tcode/{code}/layouts/{lid}/set-default")
def set_default_layout(code:str,lid:int,req:Request):
    need_admin(req)
    with closing(get_db()) as d: d.execute("UPDATE tcode_layouts SET is_default=0 WHERE tcode_code=?",(code,)); d.execute("UPDATE tcode_layouts SET is_default=1,updated_at=? WHERE id=?",(now(),lid)); d.commit()
    return {"ok":True}

@app.get("/api/tcode/{code}/schema")
def get_schema(code:str,req:Request):
    need(req)
    with closing(get_db()) as d:
        r=d.execute("SELECT * FROM tcode_schemas WHERE code=?",(code,)).fetchone()
        if not r: return {"code":code,"fields":[],"layout":{"columns":3}}
        return {"code":r["code"],"fields":json.loads(r["fields"]),"layout":json.loads(r["layout"])}

@app.post("/api/tcode/{code}/schema")
async def save_schema(code:str,req:Request):
    need_admin(req); b=await req.json()
    with closing(get_db()) as d:
        d.execute("INSERT OR REPLACE INTO tcode_schemas VALUES(?,?,?,?,?)",(code,json.dumps(b["fields"]),json.dumps(b.get("layout",{"columns":3})),now(),need(req)["username"]))
        d.execute("UPDATE tcode_layouts SET fields=?,layout_config=?,updated_at=? WHERE tcode_code=? AND is_default=1",(json.dumps(b["fields"]),json.dumps(b.get("layout",{"columns":3})),now(),code))
        d.commit()
    return {"ok":True}

@app.get("/api/tcode/{code}/records")
def list_records(code:str,req:Request):
    need(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id DESC",(code,)).fetchall()]
        for r in rows: r["data"]=json.loads(r["data"])
        return rows

@app.post("/api/tcode/{code}/records")
async def create_record(code:str,req:Request):
    u=need(req); b=await req.json()
    with closing(get_db()) as d: d.execute("INSERT INTO tcode_records(tcode_code,data,status,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(code,json.dumps(b),"active",u["username"],u["full_name"],now(),now())); d.commit(); return {"ok":True}

@app.put("/api/tcode/{code}/records/{rid}")
async def update_record(code:str,rid:int,req:Request):
    u=need(req); b=await req.json()
    with closing(get_db()) as d:
        r=d.execute("SELECT * FROM tcode_records WHERE id=?",(rid,)).fetchone()
        if r["created_by"]!=u["username"] and u["role"]!="admin": raise HTTPException(403,"not owner")
        d.execute("UPDATE tcode_records SET data=?,updated_at=? WHERE id=?",(json.dumps(b),now(),rid)); d.commit()
    return {"ok":True}

@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code:str,rid:int,req:Request):
    need_admin(req)
    with closing(get_db()) as d: d.execute("UPDATE tcode_records SET status='deleted' WHERE id=?",(rid,)); d.commit()
    return {"ok":True}

# ACCESS GRANTS
@app.get("/api/admin/access-grants")
def list_access_grants(req:Request):
    need_admin(req)
    with closing(get_db()) as d: return [dict(r) for r in d.execute("SELECT ag.*,u.full_name,u.department FROM access_grants ag LEFT JOIN users u ON ag.username=u.username ORDER BY ag.username,ag.tcode_code").fetchall()]

@app.post("/api/admin/grant-access")
def grant_access(req:Request,username:str=Form(...),tcode_code:str=Form(...),layout_code:str=Form(...)):
    u=need_admin(req)
    with closing(get_db()) as d:
        d.execute("INSERT OR REPLACE INTO access_grants(username,tcode_code,layout_code,granted_by,granted_at) VALUES(?,?,?,?,?)",(username.lower().strip(),tcode_code.upper().strip(),layout_code.upper().strip(),u["username"],now())); d.commit()
    return {"ok":True}

@app.post("/api/admin/revoke-access")
def revoke_access(req:Request,grant_id:int=Form(...)):
    need_admin(req)
    with closing(get_db()) as d: d.execute("DELETE FROM access_grants WHERE id=?",(grant_id,)); d.commit()
    return {"ok":True}

# EXCEL
@app.post("/api/admin/upload-excel-preview")
def upload_excel_preview(req:Request,file:UploadFile=File(...)):
    need_admin(req); data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True))
    headers=[str(h or "").strip() for h in rows[0]]; preview=[[str(c or "") for c in r] for r in rows[1:11] if any(r)]
    return {"headers":headers,"preview":preview,"total_rows":len([r for r in rows[1:] if any(r)])}

@app.post("/api/admin/upload-excel-save-file")
def upload_excel_save_file(req:Request,file:UploadFile=File(...),tcode_code:str=Form(...),mapping:str=Form(...)):
    u=need_admin(req); mapping_dict=json.loads(mapping); col_to_field={}
    for k,v in mapping_dict.items():
        if str(k).isdigit(): col_to_field[int(k)]=v
    data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True)); count=0
    with closing(get_db()) as d:
        for r in rows[1:]:
            if not any(r): continue
            data_dict={}
            for col_idx, field_key in col_to_field.items():
                if col_idx < len(r): data_dict[field_key]=str(r[col_idx] or "").strip()
            if not data_dict: continue
            d.execute("INSERT INTO tcode_records(tcode_code,data,status,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(tcode_code,json.dumps(data_dict),"active",u["username"],u["full_name"],now(),now())); count+=1
        d.commit()
    return {"ok":True,"saved":count}

@app.get("/api/tcode/{code}/export-excel")
def export_excel(code:str,req:Request):
    need(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id DESC",(code,)).fetchall()]
        schema=d.execute("SELECT fields FROM tcode_schemas WHERE code=?",(code,)).fetchone()
        fields=json.loads(schema["fields"]) if schema else []
        wb=Workbook(); ws=wb.active; ws.title=code
        ws.append([f["label"] for f in fields]+["Created By","Created At"])
        for r in rows:
            data=json.loads(r["data"])
            ws.append([data.get(f["key"],"") for f in fields]+[r["created_by_name"],r["created_at"]])
        bio=io.BytesIO(); wb.save(bio); bio.seek(0)
        return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={code}_export.xlsx"})

# PLANNING LOOKUP & SUGGESTIONS
@app.get("/api/planning/lookup")
def planning_lookup(req:Request,job_no:str="",item_no:str="",sap_no:str=""):
    need(req); q=(job_no or item_no or sap_no or "").strip().lower()
    if not q: return {"found":False}
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active'").fetchall()]
        for r in rows:
            data=json.loads(r["data"])
            for k,v in data.items():
                if not v: continue
                if q in str(v).strip().lower():
                    return {"found":True,"data":data,"matched_field":k}
    return {"found":False}

@app.get("/api/planning/all-jobs")
def all_jobs(req:Request):
    need(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active' LIMIT 500").fetchall()]
        jobs=[]
        for r in rows:
            data=json.loads(r["data"])
            job=data.get("job_no") or data.get("item_no") or data.get("sap_no") or ""
            if job: jobs.append({"job_no":job,"customer":data.get("customer",""),"material":data.get("material_description") or data.get("material") or ""})
        return jobs[:200]

@app.get("/api/planning/suggested-fields")
def suggested_fields(req:Request):
    need_admin(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active' LIMIT 200").fetchall()]
        field_counts={}
        for r in rows:
            data=json.loads(r["data"])
            for k in data.keys():
                field_counts[k]=field_counts.get(k,0)+1
        # also get distinct keys as suggestions
        suggestions=[{"key":k,"label":k.replace("_"," ").title(),"count":c} for k,c in sorted(field_counts.items(), key=lambda x: -x[1])]
        return {"suggestions":suggestions, "total_pln_records": len(rows)}

HTML_PAGE = r"""<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>V11 Admin Full Control</title><script src="https://cdn.tailwindcss.com"></script><link href="https://cdnjs.cloudflare.com/ajax/libs/remixicon/4.2.0/remixicon.min.css" rel="stylesheet"><style>@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');body{font-family:Inter;background:#f3f4f6}</style></head><body><div id="root"></div><div id="toasts" class="fixed top-3 right-3 z-[9999]"></div><div id="ctxMenu" class="hidden"></div><script>
const $=id=>document.getElementById(id);const toast=(m,t="ok")=>{let e=document.createElement("div");e.className=`px-4 py-3 rounded-xl text-white text-sm font-semibold ${t=="err"?"bg-red-600":"bg-gray-900"}`;e.textContent=m;$("toasts").appendChild(e);setTimeout(()=>e.remove(),3000)};let U=null,ORG=null;async function api(p,o={}){let r=await fetch(p,{credentials:"include",...o});if(!r.ok)throw new Error(await r.text());let ct=r.headers.get("content-type");if(ct&&ct.includes("application/json"))return r.json();return r;}
async function boot(){try{U=await api("/api/me");ORG=await api("/api/org");app()}catch{auth()}}boot();
function auth(){root.innerHTML=`<div class="min-h-screen flex items-center justify-center p-4 bg-[#0f172a]"><div class="bg-white rounded-2xl p-8 w-full max-w-md shadow-2xl"><h2 class="text-2xl font-bold">ProTransfer V11 - Admin Control</h2><div class="text-[11px] text-gray-500 mt-1">No Registration • Admin creates users only</div><div class="mt-6"><input id="lu" placeholder="Username" class="w-full border p-3 rounded-xl mb-3"><input id="lp" type="password" placeholder="Password" class="w-full border p-3 rounded-xl mb-4"><button onclick="login()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">Login</button><div class="mt-3 text-[11px] text-gray-500 bg-gray-50 p-2 rounded-xl">Default: admin / Admin@123</div></div></div></div>`}
async function login(){let f=new FormData();f.append("username",$("lu").value);f.append("password",$("lp").value);try{await api("/api/login",{method:"POST",body:f});location.reload()}catch(e){toast(e.message,"err")}}
function shell(inner){return `<div class="min-h-screen"><div class="bg-white border-b h-[56px] flex items-center justify-center relative"><div class="flex items-center gap-2"><i class="ri-building-4-line"></i><span class="font-bold" id="orgName">${ORG.name}</span>${U.role=='admin'?`<button onclick="editOrg()" class="w-6 h-6 rounded-full bg-gray-100 grid place-items-center"><i class="ri-pencil-line text-xs"></i></button>`:''}</div></div><div class="bg-[#f9fafb] border-b h-[52px] flex items-center justify-between px-6"><div><div class="text-[18px] font-bold">Admin Full Control Panel</div><div class="text-[11px] text-gray-500">Edit Users • Plants • Departments • Create T-Codes • Planning fields suggestion</div></div><div class="text-right"><div class="text-[14px] font-semibold">${U.full_name}</div><div class="text-[11px] text-gray-500">${U.department} • ${U.role}</div></div></div><div class="bg-white border-b h-[56px] flex items-center justify-between px-4"><div class="flex items-center gap-2"><input id="globalSearch" oninput="filterTcodes(this.value)" placeholder="Search T-Code / Layout / Job No..." class="px-4 py-2 rounded-full bg-gray-100 text-sm w-[360px]"></div><div class="flex gap-2 flex-wrap">${U.role=='admin'?`<button onclick="openAdminPanel()" class="px-3 py-2 rounded-full bg-black text-white text-xs"><i class="ri-settings-3-line"></i> Admin Panel</button><button onclick="openUserManager()" class="px-3 py-2 rounded-full bg-gray-900 text-white text-xs"><i class="ri-user-add-line"></i> Users</button><button onclick="openAccessManager()" class="px-3 py-2 rounded-full bg-indigo-600 text-white text-xs">Access</button><button onclick="openExcelUpload('PLN0')" class="px-3 py-2 rounded-full bg-green-600 text-white text-xs">Excel</button>`:''}<button onclick="v('dash')" class="px-3 py-2 rounded-full bg-gray-100 text-xs">Dashboard</button><button onclick="logout()" class="px-3 py-2 rounded-full bg-white border text-xs">Logout</button></div></div><div class="max-w-[1480px] mx-auto p-5">${inner}</div></div>`}
function app(){v('dash')}
async function v(n){if(n=='dash'){let tcodes=await api("/api/tcodes");$("root").innerHTML=shell(`<div class="flex justify-between mb-4"><h2 class="font-bold">T-Codes • ${tcodes.length} • Sub T-Codes = Layouts • Planning Excel Master</h2><span class="text-xs bg-black text-white px-3 py-1 rounded-full">V11 Admin Full Control</span></div><div id="tcodeGrid" class="grid grid-cols-4 gap-4">${tcodes.map(t=>`<div onclick="openTcode('${t.code}')" class="bg-white rounded-[20px] border p-5 cursor-pointer hover:shadow-xl"><div class="flex justify-between"><div class="w-12 h-12 rounded-xl grid place-items-center text-white" style="background:${t.color}"><i class="${t.icon}"></i></div><span class="text-[10px] px-2 py-1 rounded-full bg-gray-100">${t.layout_count} layouts</span></div><div class="mt-4"><div class="text-[11px] font-bold text-gray-500">${t.code}</div><div class="text-[14px] font-bold">${t.title}</div><div class="text-[12px] text-gray-500">${t.description}</div><div class="text-[10px] text-gray-400 mt-2">${t.allowed_depts}</div></div></div>`).join('')}</div>`)}}
function filterTcodes(q){document.querySelectorAll("#tcodeGrid > div").forEach(c=>{c.style.display=c.innerText.toLowerCase().includes(q.toLowerCase())?"":"none"})}
async function editOrg(){let nn=prompt("Org Name",ORG.name);if(!nn)return;let f=new FormData();f.append("name",nn);await api("/api/org",{method:"POST",body:f});ORG.name=nn;$("orgName").innerText=nn;toast("Updated")}

// ADMIN PANEL FULL
async function openAdminPanel(){
let depts=await api("/api/admin/departments");let plants=await api("/api/admin/plants");let tcodes=await api("/api/tcodes/all");let suggested=await api("/api/planning/suggested-fields");
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg"><i class="ri-settings-3-line"></i> Admin Panel - Edit Plants, Departments, T-Codes, Suggested Fields from Planning</h2><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div>
<div class="mt-6 grid grid-cols-12 gap-6">
<div class="col-span-6">
<div class="bg-gray-50 rounded-2xl p-5 border mb-6"><h3 class="font-bold text-sm flex justify-between"><span>Plants • ${plants.length}</span><button onclick="addPlant()" class="px-3 py-1 rounded-full bg-black text-white text-xs">+ Add</button></h3><div class="mt-3 space-y-2 max-h-[30vh] overflow-auto">${plants.map(p=>`<div class="flex items-center justify-between p-2 bg-white border rounded-xl text-xs"><span class="font-mono font-bold">${p.name} ${p.active?'<span class="text-green-600">●</span>':'<span class="text-red-600">● Inactive</span>'}</span><div class="flex gap-1"><button onclick="togglePlant('${p.name}')" class="px-2 py-1 rounded-full bg-gray-100">${p.active?'Disable':'Enable'}</button><button onclick="deletePlant('${p.name}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600">Del</button></div></div>`).join('')}</div></div>
<div class="bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm flex justify-between"><span>Departments • ${depts.length}</span><button onclick="addDept()" class="px-3 py-1 rounded-full bg-black text-white text-xs">+ Add</button></h3><div class="mt-3 space-y-2 max-h-[30vh] overflow-auto">${depts.map(d=>`<div class="flex items-center justify-between p-2 bg-white border rounded-xl text-xs"><span class="font-mono font-bold">${d.name} ${d.active?'<span class="text-green-600">●</span>':'<span class="text-red-600">● Inactive</span>'}</span><div class="flex gap-1"><button onclick="toggleDept('${d.name}')" class="px-2 py-1 rounded-full bg-gray-100">${d.active?'Disable':'Enable'}</button><button onclick="deleteDept('${d.name}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600">Del</button></div></div>`).join('')}</div></div>
</div>
<div class="col-span-6">
<div class="bg-gray-50 rounded-2xl p-5 border mb-6"><h3 class="font-bold text-sm flex justify-between"><span>T-Codes • ${tcodes.length} <span class="text-[11px] font-normal">Click to edit, or Create New</span></span><button onclick="createNewTcode()" class="px-3 py-1 rounded-full bg-black text-white text-xs">+ New T-Code</button></h3><div class="mt-3 space-y-2 max-h-[40vh] overflow-auto">${tcodes.map(t=>`<div class="flex items-center justify-between p-2 bg-white border rounded-xl text-xs"><div><div class="font-mono font-bold">${t.code} - ${t.title}</div><div class="text-[10px] text-gray-500">${t.description} • ${t.allowed_depts} • ${t.layout_count} layouts</div></div><div class="flex gap-1"><button onclick="editTcode('${t.code}')" class="px-2 py-1 rounded-full bg-gray-100">Edit</button><button onclick="deleteTcode('${t.code}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600">Del</button></div></div>`).join('')}</div></div>
<div class="bg-indigo-50 rounded-2xl p-5 border border-indigo-200"><h3 class="font-bold text-sm">Suggested Fields from Planning (PLN0 - ${suggested.total_pln_records} records)</h3><div class="text-[11px] text-gray-600 mt-1">When you create new T-Code/Layout, use these fields so auto-fill works</div><div class="mt-3 flex flex-wrap gap-2 max-h-[25vh] overflow-auto">${suggested.suggestions.map(s=>`<span class="px-2 py-1 bg-white border rounded-full text-[11px] font-mono cursor-pointer" onclick="copyField('${s.key}')">${s.key} (${s.count})</span>`).join('')||'<span class="text-xs text-gray-400">No planning data yet - Upload Excel first</span>'}</div><div class="mt-3 text-[11px] text-gray-500">Click field to copy key. Then use in Create Layout → Add Field.</div></div>
</div>
</div></div>`);
}
function copyField(k){navigator.clipboard.writeText(k);toast("Copied field key: "+k)}
async function addPlant(){let n=prompt("Plant Name e.g. Gooty Plant 2");if(!n)return;let f=new FormData();f.append("name",n);await api("/api/admin/plants/create",{method:"POST",body:f});toast("Plant added");openAdminPanel()}
async function deletePlant(n){if(!confirm("Delete plant "+n+"?"))return;let f=new FormData();f.append("name",n);await api("/api/admin/plants/delete",{method:"POST",body:f});toast("Deleted");openAdminPanel()}
async function togglePlant(n){let f=new FormData();f.append("name",n);await api("/api/admin/plants/toggle",{method:"POST",body:f});openAdminPanel()}
async function addDept(){let n=prompt("Department Name e.g. Dispatch");if(!n)return;let f=new FormData();f.append("name",n);await api("/api/admin/departments/create",{method:"POST",body:f});toast("Dept added");openAdminPanel()}
async function deleteDept(n){if(!confirm("Delete department "+n+"?"))return;let f=new FormData();f.append("name",n);await api("/api/admin/departments/delete",{method:"POST",body:f});toast("Deleted");openAdminPanel()}
async function toggleDept(n){let f=new FormData();f.append("name",n);await api("/api/admin/departments/toggle",{method:"POST",body:f});openAdminPanel()}

async function createNewTcode(){
let code=prompt("New T-Code CODE e.g. DSP0, 4 chars, UPPERCASE");if(!code)return;code=code.toUpperCase().trim();
let title=prompt("Title e.g. Dispatch Records");if(!title)return;
let desc=prompt("Description e.g. For dispatch team")||"";
let color=prompt("Color hex e.g. #ff0000", "#3b82f6")||"#3b82f6";
let icon=prompt("Icon e.g. ri-truck-line", "ri-apps-line")||"ri-apps-line";
let allowed=prompt("Allowed Departments comma separated e.g. Store,Production,Planning","Planning,Production")||"Planning";
let f=new FormData();f.append("code",code);f.append("title",title);f.append("description",desc);f.append("color",color);f.append("icon",icon);f.append("allowed_depts",allowed);
try{await api("/api/admin/tcodes/create",{method:"POST",body:f});toast("T-Code Created "+code);openAdminPanel()}catch(e){toast(e.message,"err")}
}
async function editTcode(code){
let tcodes=await api("/api/tcodes/all");let t=tcodes.find(x=>x.code==code);if(!t)return;
let title=prompt("Edit Title", t.title);if(title===null)return;
let desc=prompt("Edit Description", t.description)||t.description;
let color=prompt("Edit Color", t.color)||t.color;
let icon=prompt("Edit Icon", t.icon)||t.icon;
let allowed=prompt("Allowed Depts comma", t.allowed_depts)||t.allowed_depts;
let active=prompt("Active 1/0", t.active)||"1";
let f=new FormData();f.append("code",code);f.append("title",title);f.append("description",desc);f.append("color",color);f.append("icon",icon);f.append("allowed_depts",allowed);f.append("active",active);
await api("/api/admin/tcodes/update",{method:"POST",body:f});toast("Updated");openAdminPanel();
}
async function deleteTcode(code){if(!confirm("Delete T-Code "+code+"? All its records and layouts will be deleted."))return;let f=new FormData();f.append("code",code);try{await api("/api/admin/tcodes/delete",{method:"POST",body:f});toast("Deleted "+code);openAdminPanel()}catch(e){toast(e.message,"err")}}

// USER MANAGER WITH EDIT
async function openUserManager(){
let users=await api("/api/admin/users");let depts=await api("/api/departments");let plants=await api("/api/plants");
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg">Manage Users - Edit / Modify / Create (Admin Only)</h2><div class="flex gap-2"><button onclick="openAdminPanel()" class="px-4 py-2 rounded-full bg-black text-white text-sm">Admin Panel</button><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div></div>
<div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-5 bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Create New User</h3><div class="mt-4 space-y-3">
<input id="new_username" placeholder="Username e.g. ramesh" class="w-full border p-2.5 rounded-xl text-sm"><input id="new_password" placeholder="Password" class="w-full border p-2.5 rounded-xl text-sm"><input id="new_fullname" placeholder="Full Name" class="w-full border p-2.5 rounded-xl text-sm">
<select id="new_dept" class="w-full border p-2.5 rounded-xl text-sm">${depts.map(d=>`<option value="${d.name}">${d.name}</option>`).join('')}</select>
<select id="new_plant" class="w-full border p-2.5 rounded-xl text-sm">${plants.map(p=>`<option value="${p.name}">${p.name}</option>`).join('')}</select>
<select id="new_role" class="w-full border p-2.5 rounded-xl text-sm"><option value="user">User</option><option value="admin">Admin</option></select>
<button onclick="doCreateUser()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">+ Create User</button>
</div>
<div id="editBox" class="mt-6 hidden bg-white rounded-2xl p-4 border"><h3 class="font-bold text-sm">Edit User - <span id="editUName"></span></h3><div class="mt-3 space-y-2">
<input id="edit_fullname" class="w-full border p-2 rounded-xl text-sm"><select id="edit_dept" class="w-full border p-2 rounded-xl text-sm">${depts.map(d=>`<option value="${d.name}">${d.name}</option>`).join('')}</select><select id="edit_plant" class="w-full border p-2 rounded-xl text-sm">${plants.map(p=>`<option value="${p.name}">${p.name}</option>`).join('')}</select><select id="edit_role" class="w-full border p-2 rounded-xl text-sm"><option value="user">User</option><option value="admin">Admin</option></select><select id="edit_status" class="w-full border p-2 rounded-xl text-sm"><option value="approved">Approved</option><option value="pending">Pending</option></select>
<div class="flex gap-2"><button onclick="doUpdateUser()" class="flex-1 bg-black text-white py-2 rounded-xl text-sm">Update</button><button onclick="cancelEdit()" class="px-4 py-2 rounded-xl bg-gray-100 text-sm">Cancel</button></div></div></div>
</div>
<div class="col-span-7"><h3 class="font-bold text-sm">All Users • ${users.length}</h3><div class="mt-3 border rounded-xl overflow-auto max-h-[80vh]"><table class="w-full text-xs"><thead class="bg-gray-50 sticky top-0"><tr><th class="p-2 text-left">User</th><th class="p-2">Dept/Plant</th><th class="p-2">Role</th><th class="p-2">Action</th></tr></thead><tbody>${users.map(u=>`<tr class="border-b hover:bg-gray-50"><td class="p-2 font-mono"><b>@${u.username}</b><div class="text-[10px]">${u.full_name}</div><div class="text-[9px] text-gray-400">${u.status} • ${u.created_at.slice(0,10)}</div></td><td class="p-2">${u.department}<div class="text-[10px] text-gray-400">${u.plant}</div></td><td class="p-2"><span class="px-2 py-1 rounded-full ${u.role=='admin'?'bg-black text-white':'bg-gray-100'} text-[10px]">${u.role}</span></td><td class="p-2 flex gap-1 flex-wrap"><button onclick="startEditUser('${u.username}')" class="px-2 py-1 rounded-full bg-indigo-50 text-indigo-600 text-[11px]">Edit</button><button onclick="resetUserPass('${u.username}')" class="px-2 py-1 rounded-full bg-gray-100 text-[11px]">Reset Pass</button><button onclick="deleteUser('${u.username}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Delete</button></td></tr>`).join('')}</tbody></table></div></div></div></div>`);
}
let _editingUsername=null;
function startEditUser(username){
fetch("/api/admin/users",{credentials:"include"}).then(r=>r.json()).then(users=>{
let u=users.find(x=>x.username==username);if(!u)return;
_editingUsername=username;
$("editBox").classList.remove("hidden");$("editUName").innerText=username;
$("edit_fullname").value=u.full_name;$("edit_dept").value=u.department;$("edit_plant").value=u.plant;$("edit_role").value=u.role;$("edit_status").value=u.status;
window.scrollTo(0,0);
});
}
function cancelEdit(){_editingUsername=null;$("editBox").classList.add("hidden")}
async function doUpdateUser(){
if(!_editingUsername)return;
let f=new FormData();f.append("username",_editingUsername);f.append("full_name",$("edit_fullname").value);f.append("department",$("edit_dept").value);f.append("plant",$("edit_plant").value);f.append("role",$("edit_role").value);f.append("status",$("edit_status").value);
await api("/api/admin/update-user",{method:"POST",body:f});toast("User updated");openUserManager();
}
async function doCreateUser(){
let f=new FormData();f.append("username",$("new_username").value);f.append("password",$("new_password").value);f.append("full_name",$("new_fullname").value);f.append("department",$("new_dept").value);f.append("plant",$("new_plant").value);f.append("role",$("new_role").value);
if(!f.get("username")||!f.get("password")){toast("Username and password required","err");return}
try{await api("/api/admin/create-user",{method:"POST",body:f});toast("User created");openUserManager()}catch(e){toast(e.message,"err")}
}
async function deleteUser(username){if(!confirm("Delete @"+username+"?"))return;let f=new FormData();f.append("username",username);await api("/api/admin/delete-user",{method:"POST",body:f});toast("Deleted");openUserManager()}
async function resetUserPass(username){let np=prompt("New password for @"+username);if(!np)return;let f=new FormData();f.append("username",username);f.append("new_password",np);await api("/api/admin/reset-password",{method:"POST",body:f});toast("Password reset")}

// ACCESS MANAGER
async function openAccessManager(){
let users=await api("/api/admin/users");let tcodes=await api("/api/tcodes/all");let grants=await api("/api/admin/access-grants");
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg">Manage Access - Give T-Code / Layout to Any User</h2><div class="flex gap-2"><button onclick="openUserManager()" class="px-4 py-2 rounded-full bg-black text-white text-sm">Users</button><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div></div>
<div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-5 bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Grant Access</h3><div class="mt-4 space-y-3">
<select id="acc_user" class="w-full border p-2.5 rounded-xl text-sm"><option value="">-- Select User --</option>${users.map(u=>`<option value="${u.username}">${u.full_name} @${u.username} - ${u.department}</option>`).join('')}</select>
<select id="acc_tcode" class="w-full border p-2.5 rounded-xl text-sm" onchange="loadLayoutsForAccess(this.value)"><option value="">-- Select T-Code --</option>${tcodes.map(t=>`<option value="${t.code}">${t.code} - ${t.title}</option>`).join('')}</select>
<select id="acc_layout" class="w-full border p-2.5 rounded-xl text-sm"><option value="">-- All Layouts (Wildcard) --</option></select>
<button onclick="doGrantAccess()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">+ Grant</button>
</div></div>
<div class="col-span-7"><h3 class="font-bold text-sm">Current Grants • ${grants.length}</h3><div class="mt-3 border rounded-xl overflow-auto max-h-[70vh]"><table class="w-full text-xs"><thead class="bg-gray-50 sticky top-0"><tr><th class="p-2 text-left">User</th><th class="p-2">T-Code</th><th class="p-2">Layout</th><th class="p-2">Action</th></tr></thead><tbody>${grants.map(g=>`<tr class="border-b"><td class="p-2"><b>${g.full_name||g.username}</b><div class="text-[10px] text-gray-500">@${g.username}</div></td><td class="p-2 font-mono font-bold">${g.tcode_code}</td><td class="p-2 font-mono">${g.layout_code||'ALL'}</td><td class="p-2"><button onclick="revokeGrant(${g.id})" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Revoke</button></td></tr>`).join('')}</tbody></table></div></div></div></div>`);
}
async function loadLayoutsForAccess(tcode){let sel=$("acc_layout");sel.innerHTML='<option value="">-- All Layouts --</option>';if(!tcode)return;let layouts=await api("/api/tcode/"+tcode+"/layouts");sel.innerHTML='<option value="">-- All Layouts (Wildcard) --</option>'+layouts.map(l=>`<option value="${l.layout_code}">${l.layout_code} - ${l.title}</option>`).join('');}
async function doGrantAccess(){let user=$("acc_user").value;let tcode=$("acc_tcode").value;let layout=$("acc_layout").value;if(!user||!tcode){toast("Select User and T-Code","err");return}let f=new FormData();f.append("username",user);f.append("tcode_code",tcode);f.append("layout_code",layout);try{await api("/api/admin/grant-access",{method:"POST",body:f});toast(`Granted ${user} → ${tcode}`);openAccessManager()}catch(e){toast(e.message,"err")}}
async function revokeGrant(id){if(!confirm("Revoke?"))return;let f=new FormData();f.append("grant_id",id);await api("/api/admin/revoke-access",{method:"POST",body:f});toast("Revoked");openAccessManager()}

// TCODES VIEW WITH SUGGESTIONS
async function openTcode(code){
let layouts=await api("/api/tcode/"+code+"/layouts");let tc=(await api("/api/tcodes/all").catch(()=>api("/api/tcodes"))).find(x=>x.code==code)||{code:code,title:code,color:"#000",icon:"ri-apps-line"};let saved=localStorage.getItem("layout_"+code);let curLayout=layouts.find(l=>l.id==saved)||layouts.find(l=>l.is_default)||layouts[0];let records=await api("/api/tcode/"+code+"/records");if(!curLayout){$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6">No layouts</div>`);return}
let formHtml=curLayout.fields.map(f=>{let isJob=f.key.toLowerCase().includes("job")||f.key.toLowerCase().includes("item")||f.key.toLowerCase().includes("sap");let onBlur=isJob?`onblur="lookupPlanning('${f.key}',this.value)"`:"";return `<div class="col-span-${f.width||1}"><label class="text-[11px] font-semibold">${f.label}${f.required?' *':''} ${isJob?'<span class="text-[9px] px-1 py-0.5 rounded bg-blue-50 text-blue-600">🔍 Auto</span>':''}</label>${f.type=='textarea'?`<textarea id="f_${f.key}" ${onBlur} class="w-full border p-2.5 rounded-xl mt-1" rows="2"></textarea>`:f.type=='select'?`<select id="f_${f.key}" class="w-full border p-2.5 rounded-xl mt-1"><option value="">Select</option>${(f.options||[]).map(o=>`<option>${o}</option>`).join('')}</select>`:f.type=='date'?`<input id="f_${f.key}" type="date" ${onBlur} class="w-full border p-2.5 rounded-xl mt-1">`:f.type=='number'?`<input id="f_${f.key}" type="number" class="w-full border p-2.5 rounded-xl mt-1">`:`<input id="f_${f.key}" type="text" ${onBlur} class="w-full border p-2.5 rounded-xl mt-1" list="jobList">`}<div id="hint_${f.key}" class="text-[10px] text-green-600 mt-1 hidden"></div></div>`}).join('');formHtml+=`<datalist id="jobList"></datalist>`;
$("root").innerHTML=shell(`
<div class="flex justify-between items-center mb-4"><div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl grid place-items-center text-white" style="background:${tc.color}"><i class="${tc.icon}"></i></div><div><div class="text-[11px] font-bold text-gray-500">${tc.code} • ${layouts.length} sub T-Codes • PLN0 linked</div><div class="text-[18px] font-bold">${tc.title}</div></div></div><div class="flex gap-2"><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-white border text-sm">← Dashboard</button>${U.role=='admin'?`<button onclick="openExcelUpload('${code}')" class="px-4 py-2 rounded-full bg-green-600 text-white text-sm">Excel</button><button onclick="editSchema('${code}')" class="px-4 py-2 rounded-full bg-indigo-600 text-white text-sm">Edit Form + Suggestions</button>`:''}</div></div>
<div class="bg-white rounded-2xl border p-4 mb-4 flex items-center gap-4 flex-wrap"><label class="text-sm font-bold"><i class="ri-layout-line"></i> Choose Layout:</label><select id="layoutSelect" onchange="switchLayout('${code}',this.value)" class="flex-1 border p-2.5 rounded-xl text-sm min-w-[340px]">${layouts.map(l=>`<option value="${l.id}" ${l.id==curLayout.id?'selected':''}>${l.layout_code} - ${l.title} ${l.is_default?' (Default)':''}</option>`).join('')}</select><button onclick="showLayoutDetails()" class="px-3 py-2 rounded-full bg-gray-100 text-sm">Details</button><button onclick="openLayoutMenu(event)" class="px-3 py-2 rounded-full bg-gray-100 text-sm">⋮ Options</button><a href="/api/tcode/${code}/export-excel" class="px-3 py-2 rounded-full bg-gray-100 text-sm">Export</a></div>
<div class="grid grid-cols-12 gap-5"><div class="col-span-5 bg-white rounded-2xl border p-5" oncontextmenu="openCtx(event,'${code}')"><h3 class="font-bold text-sm flex justify-between"><span>${curLayout.layout_code} : ${curLayout.title}</span><span class="text-[10px] px-2 py-1 rounded-full ${curLayout.is_default?'bg-green-100 text-green-700':'bg-gray-100'}">${curLayout.is_default?'★ Default':'Custom'}</span></h3><div id="formArea" class="mt-4 grid grid-cols-3 gap-3">${formHtml}</div><div class="mt-4 flex gap-2"><button onclick="saveRecord('${code}')" class="flex-1 bg-black text-white py-3 rounded-xl">+ Save</button><button onclick="clearForm()" class="px-5 py-3 rounded-xl bg-gray-100">Clear</button></div><div id="editHint" class="text-xs text-gray-400 mt-2"></div></div><div class="col-span-7 bg-white rounded-2xl border p-5"><div class="flex justify-between"><h3 class="font-bold text-sm">Records • ${records.length}</h3><input id="recSearch" oninput="filterRecords(this.value)" placeholder="Search..." class="border px-3 py-1 rounded-full text-sm"></div><div class="mt-3 overflow-auto max-h-[70vh] border rounded-xl"><table class="w-full text-xs"><thead class="bg-gray-50"><tr><th class="p-2">ID</th><th class="p-2">Layout</th><th class="p-2">Data</th><th class="p-2">By</th><th class="p-2">Action</th></tr></thead><tbody id="recBody">${records.map(r=>`<tr class="border-b"><td class="p-2">#${r.id}</td><td class="p-2 font-mono text-[11px]">${curLayout.layout_code}</td><td class="p-2">${Object.entries(r.data).slice(0,3).map(([k,v])=>`<span class="mr-2"><b>${k}:</b> ${String(v).slice(0,20)}</span>`).join('')}</td><td class="p-2">${r.created_by_name}<div class="text-[10px] text-gray-400">${r.created_at.slice(0,16).replace('T',' ')}</div></td><td class="p-2"><button onclick="loadEdit(${r.id})" class="px-2 py-1 rounded-full bg-gray-100 text-[11px]">Edit</button>${U.role=='admin'?`<button onclick="delRecord('${code}',${r.id})" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Del</button>`:''}</td></tr>`).join('')}</tbody></table></div></div></div>`);
window._curLayouts=layouts;window._curLayout=curLayout;window._curRecords=records;window._curCode=code;setTimeout(loadJobList,500);
}
function switchLayout(code,id){localStorage.setItem("layout_"+code,id);openTcode(code)}
async function saveAsLayout(code){let layout_code=prompt("Enter Layout Code e.g. "+code+"-QC02","");let title=prompt("Enter Title for this Sub T-Code","");if(!title){toast("Title required","err");return}let desc=prompt("Description","")||"";let fields=window._curLayout.fields;let layout_config=window._curLayout.layout_config||{columns:3};try{let res=await api("/api/tcode/"+code+"/layouts",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({layout_code:layout_code,name:title,title:title,fields:fields,layout_config:layout_config,description:desc})});toast("Saved "+res.layout_code);openTcode(code);}catch(e){toast(e.message,"err")}}
function showLayoutDetails(){let l=window._curLayout;let html=`<div class="bg-white rounded-2xl border p-5 max-w-md"><h3 class="font-bold">${l.layout_code} : ${l.title}</h3><div class="mt-4 space-y-2 text-sm"><div class="flex justify-between"><span class="text-gray-500">Layout Code</span><b class="font-mono">${l.layout_code}</b></div><div class="flex justify-between"><span class="text-gray-500">Created By</span><b>${l.created_by_name} @${l.created_by}</b></div><div class="flex justify-between"><span class="text-gray-500">Created</span><b>${new Date(l.created_at).toLocaleString()}</b></div><div class="flex justify-between"><span class="text-gray-500">Modified</span><b>${new Date(l.updated_at).toLocaleString()}</b></div></div><button onclick="hideCtx()" class="mt-4 w-full bg-black text-white py-2 rounded-xl">Close</button></div>`;let ctx=$("ctxMenu");ctx.innerHTML=html;ctx.className="fixed inset-0 bg-black/40 grid place-items-center z-[9999]";ctx.classList.remove("hidden");}
function openLayoutMenu(e){e.stopPropagation();let l=window._curLayout;let code=window._curCode;let html=`<div style="position:fixed;left:${e.clientX}px;top:${e.clientY}px;background:white;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 20px 40px rgba(0,0,0,.15);padding:6px;min-width:280px"><div class="p-2 space-y-1"><div class="px-3 py-2 font-bold text-sm">${l.layout_code} : ${l.title}</div><button onclick="showLayoutDetails();hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm">Details</button><button onclick="saveAsLayout('${code}');hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm">Save as New Sub T-Code</button><button onclick="setDefaultLayout('${code}');hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm">Set Default</button><button onclick="deleteLayout('${code}',${l.id});hideCtx()" class="w-full text-left px-3 py-2 hover:bg-red-50 text-red-600 rounded-xl text-sm">Delete</button></div></div>`;let ctx=$("ctxMenu");ctx.innerHTML=html;ctx.className="fixed inset-0 z-[9999]";ctx.classList.remove("hidden");}
function openCtx(e,code){e.preventDefault();let l=window._curLayout;let html=`<div style="position:fixed;left:${e.clientX}px;top:${e.clientY}px;background:white;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 20px 40px rgba(0,0,0,.15);padding:8px;min-width:300px"><div class="px-3 py-2 font-bold text-sm">${l.layout_code} : ${l.title} ${l.is_default?'★':''}</div><div class="px-3 py-1 text-[11px] text-gray-500">Created By: ${l.created_by_name} @${l.created_by}</div><div class="px-3 py-1 text-[11px] text-gray-500">Created: ${new Date(l.created_at).toLocaleString()}</div><div class="px-3 py-1 text-[11px] text-gray-500">Modified: ${new Date(l.updated_at).toLocaleString()}</div></div>`;let ctx=$("ctxMenu");ctx.innerHTML=html;ctx.className="fixed inset-0 z-[9999]";ctx.classList.remove("hidden");}
function hideCtx(){$("ctxMenu").classList.add("hidden");$("ctxMenu").innerHTML=""}
document.addEventListener("click",()=>{if(!$("ctxMenu").classList.contains("hidden"))hideCtx()})
async function setDefaultLayout(code){if(U.role!='admin'){toast("Admin only","err");return}await api("/api/tcode/"+code+"/layouts/"+window._curLayout.id+"/set-default",{method:"POST"});toast("Set default");openTcode(code)}
async function deleteLayout(code,id){if(window._curLayout.is_default){toast("Cannot delete default","err");return}if(!confirm("Delete?"))return;await api("/api/tcode/"+code+"/layouts/"+id,{method:"DELETE"});localStorage.removeItem("layout_"+code);toast("Deleted");openTcode(code)}
let _editId=null;
function clearForm(){_editId=null;document.querySelectorAll("#formArea input, #formArea textarea, #formArea select").forEach(e=>e.value="");$("editHint").innerText=""}
function loadEdit(id){let rec=window._curRecords.find(x=>x.id==id);if(!rec)return;_editId=id;for(let k in rec.data){let el=document.getElementById("f_"+k);if(el)el.value=rec.data[k]}$("editHint").innerText="Editing #"+id;window.scrollTo(0,0)}
async function saveRecord(code){let data={};for(let f of window._curLayout.fields){let el=document.getElementById("f_"+f.key);if(!el)continue;if(f.required&&!el.value){toast(f.label+" required","err");return}data[f.key]=el.value}try{if(_editId){await api("/api/tcode/"+code+"/records/"+_editId,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});toast("Updated")}else{await api("/api/tcode/"+code+"/records",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});toast("Saved")}clearForm();openTcode(code)}catch(e){toast(e.message,"err")}}
async function delRecord(code,id){if(!confirm("Delete #"+id+"?"))return;await api("/api/tcode/"+code+"/records/"+id,{method:"DELETE"});toast("Deleted");openTcode(code)}
function filterRecords(q){document.querySelectorAll("#recBody tr").forEach(r=>{r.style.display=r.innerText.toLowerCase().includes(q.toLowerCase())?"":"none"})}

async function editSchema(code){
let schema=await api("/api/tcode/"+code+"/schema");
let suggested=await api("/api/planning/suggested-fields");
let fields=schema.fields;
let listHtml=fields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label} (${f.type})<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('');
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between flex-wrap gap-2"><h2 class="font-bold">Edit Form - ${code} + Suggestions from Planning (${suggested.total_pln_records} PLN records)</h2><div class="flex gap-2"><button onclick="openTcode('${code}')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button><button onclick="saveSchema('${code}')" class="px-4 py-2 rounded-full bg-black text-white text-sm">Save</button></div></div>
<div class="mt-4 bg-indigo-50 border border-indigo-200 rounded-2xl p-4"><h3 class="font-bold text-sm">💡 Suggested Fields from Planning Excel - Click to Add</h3><div class="text-[11px] text-gray-600">These are fields from your uploaded PLN0 Excel. Adding them enables auto-fill by Job No.</div><div class="mt-2 flex flex-wrap gap-2">${suggested.suggestions.map(s=>`<button onclick="quickAddSuggested('${s.key}','${s.label}')" class="px-3 py-1.5 bg-white border border-indigo-200 rounded-full text-[11px] hover:bg-indigo-100">+ ${s.label} (${s.key}) • ${s.count}</button>`).join('')||'<span class="text-xs text-gray-400">Upload Planning Excel first to see suggestions</span>'}</div></div>
<div class="mt-6 grid grid-cols-12 gap-5"><div class="col-span-7"><div id="fieldList" class="space-y-2">${listHtml}</div></div><div class="col-span-5 bg-gray-50 rounded-2xl p-4 border"><h3 class="font-semibold text-sm">Add Field Manually</h3><div class="mt-3 space-y-3"><input id="nf_key" placeholder="key e.g. job_no" class="w-full border p-2.5 rounded-xl text-sm"><input id="nf_label" placeholder="Label e.g. Job No" class="w-full border p-2.5 rounded-xl text-sm"><select id="nf_type" class="w-full border p-2.5 rounded-xl text-sm"><option value="text">Text</option><option value="number">Number</option><option value="date">Date</option><option value="textarea">Long Text</option><option value="select">Dropdown</option></select><input id="nf_options" placeholder="Options comma e.g. Pass,Fail" class="w-full border p-2.5 rounded-xl text-sm"><div class="grid grid-cols-2 gap-3"><select id="nf_width" class="border p-2.5 rounded-xl text-sm"><option value="1">W1</option><option value="2">W2</option><option value="3">W Full</option></select><label class="flex items-center gap-2 text-sm"><input id="nf_req" type="checkbox"> Required</label></div><button onclick="addField()" class="w-full bg-black text-white py-2.5 rounded-xl text-sm">+ Add Field</button></div></div></div></div>`);window._editFields=fields
}
function quickAddSuggested(key,label){
let field={key:key,label:label,type:"text",width:1,required:false};
window._editFields.push(field);
$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('');toast("Added "+key+" from Planning suggestion")
}
function addField(){let key=$("nf_key").value.trim().replace(/\s+/g,'_').toLowerCase();if(!key){toast("Enter key","err");return}let field={key:key,label:$("nf_label").value||key,type:$("nf_type").value,width:parseInt($("nf_width").value),required:$("nf_req").checked};if(field.type=='select'){field.options=$("nf_options").value.split(',').map(s=>s.trim()).filter(Boolean)}window._editFields.push(field);$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('')}
function removeField(i){window._editFields.splice(i,1);$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('')}
function moveField(i,dir){let j=i+dir;if(j<0||j>=window._editFields.length)return;let t=window._editFields[i];window._editFields[i]=window._editFields[j];window._editFields[j]=t;$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('')}
async function saveSchema(code){await api("/api/tcode/"+code+"/schema",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({fields:window._editFields,layout:{columns:3}})});toast("Saved");openTcode(code)}

// EXCEL
async function openExcelUpload(code){
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6 max-w-4xl mx-auto"><div class="flex justify-between"><h2 class="font-bold text-lg">Upload Excel to ${code} - Planning Details</h2><button onclick="openTcode('${code}')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div>
<div class="mt-6 space-y-4"><div class="border-2 border-dashed rounded-2xl p-8 text-center bg-gray-50"><input id="excelFile" type="file" accept=".xlsx,.xls" class="hidden" onchange="previewExcel('${code}')"><button onclick="document.getElementById('excelFile').click()" class="px-6 py-3 bg-black text-white rounded-xl">Choose Excel</button><div id="fileName" class="mt-2 text-sm font-semibold"></div></div><div id="excelPreview" class="hidden"><h3 class="font-bold text-sm">Preview - Map Excel Columns to Fields (Suggestions from existing PLN0 fields auto-selected)</h3><div class="mt-3 overflow-auto border rounded-xl max-h-[40vh]"><table class="w-full text-xs" id="previewTable"></table></div><div id="mappingArea" class="mt-4 grid grid-cols-3 gap-3"></div><button onclick="saveExcelToPlanning('${code}')" class="mt-4 w-full bg-green-600 text-white py-3 rounded-xl font-semibold">Load All Rows to ${code}</button></div></div></div>`);
}
let _uploadFile=null;
async function previewExcel(code){
let fileEl=$("excelFile");if(!fileEl.files[0])return;$("fileName").innerText=fileEl.files[0].name;
let f=new FormData();f.append("file",fileEl.files[0]);let res=await api("/api/admin/upload-excel-preview",{method:"POST",body:f});
let fields=(await api("/api/tcode/"+code+"/schema")).fields;
let mappingHtml=fields.map(ff=>`<div><label class="text-[11px] font-semibold">${ff.label} (${ff.key})</label><select id="map_${ff.key}" class="w-full border p-2 rounded-xl text-xs mt-1"><option value="">-- Ignore --</option>${res.headers.map((h,i)=>`<option value="${i}" ${h.toLowerCase().includes(ff.key.toLowerCase())||ff.key.toLowerCase().includes(h.toLowerCase().split(' ')[0])?'selected':''}>${h} (Col ${i+1})</option>`).join('')}</select></div>`).join('');
$("mappingArea").innerHTML=mappingHtml;
let tableHtml=`<thead class="bg-gray-900 text-white"><tr>${res.headers.map(h=>`<th class="p-2 text-left">${h}</th>`).join('')}</tr></thead><tbody>${res.preview.map(r=>`<tr class="border-b">${r.map(c=>`<td class="p-2">${c}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td colspan="${res.headers.length}" class="p-2 text-center text-gray-500">Showing 10 of ${res.total_rows} rows</td></tr></tfoot>`;
$("previewTable").innerHTML=tableHtml;$("excelPreview").classList.remove("hidden");_uploadFile=fileEl.files[0];
}
async function saveExcelToPlanning(code){
if(!_uploadFile){toast("Choose file first","err");return}
let fields=(await api("/api/tcode/"+code+"/schema")).fields;let mapping={};
fields.forEach(ff=>{let sel=$("map_"+ff.key);if(sel&&sel.value!==""){mapping[sel.value]=ff.key}});
let fd2=new FormData();fd2.append("file",_uploadFile);fd2.append("tcode_code",code);fd2.append("mapping",JSON.stringify(mapping));
let saveRes=await api("/api/admin/upload-excel-save-file",{method:"POST",body:fd2});
toast(`Loaded ${saveRes.saved} rows to ${code}`);openTcode(code);
}

async function lookupPlanning(fieldKey,value){
if(!value||value.length<2)return;
try{
let res=await api("/api/planning/lookup?job_no="+encodeURIComponent(value.trim()));
if(res.found){
let data=res.data;
for(let k in data){
let el=document.getElementById("f_"+k);
if(!el){
 for(let fk of window._curLayout.fields){
  if(k.toLowerCase().includes(fk.key.toLowerCase())||fk.key.toLowerCase().includes(k.toLowerCase())){
   el=document.getElementById("f_"+fk.key);if(el)break;
  }
 }
}
if(el&&!el.value&&data[k]){el.value=data[k];el.classList.add("bg-green-50");}
}
let hintEl=$("hint_"+fieldKey);if(hintEl){hintEl.innerText=`✓ From Planning: ${res.matched_field}`;hintEl.classList.remove("hidden")}
toast("Auto-filled from Planning");
}
}catch(e){}
}
async function loadJobList(){try{let jobs=await api("/api/planning/all-jobs");let dl=$("jobList");if(!dl)return;dl.innerHTML=jobs.map(j=>`<option value="${j.job_no}">${j.customer}</option>`).join('');}catch(e){}}

async function logout(){await api("/api/logout",{method:"POST"});location.reload()}
</script></body></html>
"""

@app.get("/",response_class=HTMLResponse)
def index(): return HTML_PAGE

if __name__=="__main__":
    import uvicorn
    uvicorn.run(app,host="0.0.0.0",port=int(os.environ.get("PORT",8000)))