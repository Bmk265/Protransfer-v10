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
        CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY,password_hash TEXT,full_name TEXT,role TEXT,department TEXT,sub_department TEXT,plant TEXT,status TEXT DEFAULT 'pending',created_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,username TEXT,expires TEXT);
        CREATE TABLE IF NOT EXISTS departments(name TEXT PRIMARY KEY,active INT DEFAULT 1);
        CREATE TABLE IF NOT EXISTS org(id INT PRIMARY KEY DEFAULT 1,name TEXT DEFAULT 'TCL Pvt Ltd - Gooty Plant (1400)');
        CREATE TABLE IF NOT EXISTS tcodes(code TEXT PRIMARY KEY,title TEXT,description TEXT,color TEXT,icon TEXT,allowed_depts TEXT,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_schemas(code TEXT PRIMARY KEY,fields TEXT,layout TEXT,updated_at TEXT,updated_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_records(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,data TEXT,status TEXT DEFAULT 'active',created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_layouts(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,layout_code TEXT UNIQUE,name TEXT,title TEXT,is_default INT DEFAULT 0,fields TEXT,layout_config TEXT,description TEXT,created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS access_grants(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT,tcode_code TEXT,layout_code TEXT,granted_by TEXT,granted_at TEXT,UNIQUE(username,tcode_code,layout_code));
        CREATE TABLE IF NOT EXISTS transfers(id INTEGER PRIMARY KEY AUTOINCREMENT,transfer_number TEXT UNIQUE,sender_username TEXT,sender_name TEXT,from_dept TEXT,to_dept TEXT,target_username TEXT,target_name TEXT,from_serial REAL,to_serial REAL,row_count INT,production_name TEXT,qc_name TEXT,format_id INT,status TEXT DEFAULT 'pending',created_at TEXT);
        CREATE TABLE IF NOT EXISTS rows(id INTEGER PRIMARY KEY AUTOINCREMENT,transfer_id INT,serial_no REAL,handover_date TEXT,job_no TEXT,customer TEXT,material_description TEXT,specified_length REAL,entry_length REAL,drum_size TEXT,drum_code TEXT,batch_no TEXT,entry_made_by TEXT);
        CREATE TABLE IF NOT EXISTS formats(id INTEGER PRIMARY KEY AUTOINCREMENT,from_dept TEXT,to_dept TEXT,header_title TEXT,section_title TEXT,format_no TEXT,columns TEXT,active INT DEFAULT 1);
        """)
        for d in ["Production","Quality Control","Final QC","Commercial","Planning","Store"]: c.execute("INSERT OR IGNORE INTO departments VALUES(?,1)",(d,))
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
            c.execute("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,1,?)",(code,title,desc,color,icon,depts,datetime.datetime.now().isoformat()))
        ncr=json.dumps([{"key":"date","label":"Date","type":"date","required":True,"width":1},{"key":"job_no","label":"Job No / SAP No","type":"text","required":True,"width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"material","label":"Material Description","type":"textarea","width":2},{"key":"nc_type","label":"NC Type","type":"select","options":["Dimension","Visual","Mechanical","Other"],"width":1},{"key":"description","label":"NC Description","type":"textarea","width":3},{"key":"status","label":"Status","type":"select","options":["Open","Closed"],"width":1}])
        gch=json.dumps([{"key":"date","label":"Green Card Date","type":"date","width":1},{"key":"job_no","label":"Job No","type":"text","width":1},{"key":"drum_code","label":"Drum Code","type":"text","width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"qc_status","label":"QC Status","type":"select","options":["Pass","Hold","Reject"],"width":1},{"key":"remarks","label":"Remarks","type":"textarea","width":3}])
        pln=json.dumps([{"key":"job_no","label":"Job No","type":"text","required":True,"width":1},{"key":"item_no","label":"Item No","type":"text","width":1},{"key":"sap_no","label":"SAP No","type":"text","width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"material_description","label":"Material Description","type":"textarea","width":2},{"key":"specified_length","label":"Specified Length","type":"number","width":1},{"key":"drum_size","label":"Drum Size","type":"text","width":1},{"key":"plan_date","label":"Plan Date","type":"date","width":1},{"key":"remarks","label":"Remarks","type":"textarea","width":2}])
        c.execute("INSERT OR IGNORE INTO tcode_schemas VALUES(?,?,?,?,?)",("NCR0",ncr,json.dumps({"columns":3}),datetime.datetime.now().isoformat(),"system"))
        c.execute("INSERT OR IGNORE INTO tcode_schemas VALUES(?,?,?,?,?)",("GCH0",gch,json.dumps({"columns":3}),datetime.datetime.now().isoformat(),"system"))
        c.execute("INSERT OR IGNORE INTO tcode_schemas VALUES(?,?,?,?,?)",("PLN0",pln,json.dumps({"columns":3}),datetime.datetime.now().isoformat(),"system"))
        # Default layouts with layout_code as sub tcode
        c.execute("INSERT OR IGNORE INTO tcode_layouts(id,tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(1,'NCR0','NCR0-001','Default Layout','Default NCR View',1,?,?,?, 'system','System',?,?)",(ncr,json.dumps({"columns":3}),"System default layout for NCR0",datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        c.execute("INSERT OR IGNORE INTO tcode_layouts(id,tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(2,'GCH0','GCH0-001','Default Layout','Default GCH View',1,?,?,?, 'system','System',?,?)",(gch,json.dumps({"columns":3}),"System default for GCH0",datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        c.execute("INSERT OR IGNORE INTO tcode_layouts(id,tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(3,'PLN0','PLN0-001','Default Layout','Default Planning View',1,?,?,?, 'system','System',?,?)",(pln,json.dumps({"columns":3}),"System default - Master planning",datetime.datetime.now().isoformat(),datetime.datetime.now().isoformat()))
        # Default admin
        ph=bcrypt.hashpw("Admin@123".encode(),bcrypt.gensalt()).decode()
        c.execute("INSERT OR IGNORE INTO users VALUES(?,?,?,?,?,?,?,?,?)",("admin",ph,"System Admin","admin","Planning","","Gooty","approved",datetime.datetime.now().isoformat()))
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

@app.post("/api/admin/create-user")
def admin_create_user(req:Request,username:str=Form(...),password:str=Form(...),full_name:str=Form(...),department:str=Form(...),plant:str=Form(...),role:str=Form(...)):
    need_admin(req)
    username=username.lower().strip()
    if not username or not password: raise HTTPException(400,"Username and password required")
    with closing(get_db()) as d:
        if d.execute("SELECT 1 FROM users WHERE username=?",(username,)).fetchone(): raise HTTPException(400,f"User {username} already exists")
        ph=bcrypt.hashpw(password.encode(),bcrypt.gensalt()).decode()
        d.execute("INSERT INTO users(username,password_hash,full_name,role,department,sub_department,plant,status,created_at) VALUES(?,?,?,?,?,?,?, ?, ?)",(username,ph,full_name,role,department,"",plant,"approved",now()))
        d.commit()
    return {"ok":True}

@app.get("/api/admin/users")
def admin_list_users(req:Request):
    need_admin(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT username,full_name,role,department,plant,status,created_at FROM users ORDER BY created_at DESC").fetchall()]
        return rows

@app.post("/api/admin/delete-user")
def admin_delete_user(req:Request,username:str=Form(...)):
    need_admin(req); username=username.lower().strip()
    if username==need(req)["username"]: raise HTTPException(400,"Cannot delete yourself")
    with closing(get_db()) as d:
        d.execute("DELETE FROM users WHERE username=?",(username,))
        d.execute("DELETE FROM sessions WHERE username=?",(username,))
        d.execute("DELETE FROM access_grants WHERE username=?",(username,))
        d.commit()
    return {"ok":True}

@app.post("/api/admin/reset-password")
def admin_reset_password(req:Request,username:str=Form(...),new_password:str=Form(...)):
    need_admin(req); ph=bcrypt.hashpw(new_password.encode(),bcrypt.gensalt()).decode()
    with closing(get_db()) as d:
        d.execute("UPDATE users SET password_hash=? WHERE username=?",(ph,username.lower().strip()))
        d.commit()
    return {"ok":True}

@app.get("/api/tcodes")
def list_tcodes(req:Request):
    u=need(req)
    with closing(get_db()) as d:
        all_c=[dict(r) for r in d.execute("SELECT * FROM tcodes WHERE active=1 ORDER BY code").fetchall()]
        for tc in all_c:
            cnt=d.execute("SELECT COUNT(*) c FROM tcode_layouts WHERE tcode_code=?",(tc["code"],)).fetchone()["c"]
            tc["layout_count"]=cnt
        if u["role"]=="admin": return all_c
        grants=[dict(r) for r in d.execute("SELECT tcode_code,layout_code FROM access_grants WHERE username=?",(u["username"],)).fetchall()]
        granted_tcodes=set(g["tcode_code"] for g in grants)
        visible=[]
        for tc in all_c:
            allowed=[x.strip() for x in tc["allowed_depts"].split(",")]
            if u["department"] in allowed or tc["code"] in granted_tcodes: visible.append(tc)
        return visible

@app.get("/api/tcodes/all")
def list_tcodes_all(req:Request):
    need_admin(req)
    with closing(get_db()) as d:
        all_c=[dict(r) for r in d.execute("SELECT * FROM tcodes WHERE active=1 ORDER BY code").fetchall()]
        for tc in all_c:
            cnt=d.execute("SELECT COUNT(*) c FROM tcode_layouts WHERE tcode_code=?",(tc["code"],)).fetchone()["c"]
            tc["layout_count"]=cnt
        return all_c

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
        tcode_row=d.execute("SELECT allowed_depts FROM tcodes WHERE code=?",(code,)).fetchone()
        dept_allowed=False
        if tcode_row:
            allowed=[x.strip() for x in tcode_row["allowed_depts"].split(",")]
            if u["department"] in allowed: dept_allowed=True
        if dept_allowed: return rows
        if not grants: return rows
        has_wildcard=any(g["layout_code"]=="" for g in grants)
        if has_wildcard: return rows
        allowed_layout_codes=set(g["layout_code"] for g in grants if g["layout_code"])
        return [r for r in rows if r["layout_code"] in allowed_layout_codes]

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
            raise HTTPException(400,f"Layout Code {layout_code} already exists")
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
    with closing(get_db()) as d:
        d.execute("UPDATE tcode_layouts SET is_default=0 WHERE tcode_code=?",(code,))
        d.execute("UPDATE tcode_layouts SET is_default=1,updated_at=? WHERE id=?",(now(),lid))
        d.commit()
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
    with closing(get_db()) as d:
        d.execute("INSERT INTO tcode_records(tcode_code,data,status,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(code,json.dumps(b),"active",u["username"],u["full_name"],now(),now())); d.commit()
        return {"ok":True}

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
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT ag.*,u.full_name,u.department FROM access_grants ag LEFT JOIN users u ON ag.username=u.username ORDER BY ag.username,ag.tcode_code").fetchall()]
        return rows

@app.post("/api/admin/grant-access")
def grant_access(req:Request,username:str=Form(...),tcode_code:str=Form(...),layout_code:str=Form(...)):
    u=need_admin(req); username=username.lower().strip(); tcode_code=tcode_code.upper().strip(); layout_code=layout_code.upper().strip() if layout_code else ""
    with closing(get_db()) as d:
        if not d.execute("SELECT 1 FROM users WHERE username=?",(username,)).fetchone(): raise HTTPException(404,"User not found")
        d.execute("INSERT OR REPLACE INTO access_grants(username,tcode_code,layout_code,granted_by,granted_at) VALUES(?,?,?,?,?)",(username,tcode_code,layout_code,u["username"],now()))
        d.commit()
    return {"ok":True}

@app.post("/api/admin/revoke-access")
def revoke_access(req:Request,grant_id:int=Form(...)):
    need_admin(req)
    with closing(get_db()) as d: d.execute("DELETE FROM access_grants WHERE id=?",(grant_id,)); d.commit()
    return {"ok":True}

# EXCEL UPLOAD
@app.post("/api/admin/upload-excel-preview")
def upload_excel_preview(req:Request,file:UploadFile=File(...)):
    need_admin(req); data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True))
    if not rows: raise HTTPException(400,"Empty sheet")
    headers=[str(h or "").strip() for h in rows[0]]; preview_rows=[]
    for r in rows[1:11]:
        if not any(r): continue
        preview_rows.append([str(c or "") for c in r])
    all_count=len([r for r in rows[1:] if any(r)])
    return {"headers":headers,"preview":preview_rows,"total_rows":all_count}

@app.post("/api/admin/upload-excel-preview-full")
def upload_excel_preview_full(req:Request,file:UploadFile=File(...)):
    need_admin(req); data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True))
    if not rows: raise HTTPException(400,"Empty")
    headers=[str(h or "").strip() for h in rows[0]]; all_rows=[]
    for r in rows[1:]:
        if not any(r): continue
        all_rows.append([str(c or "") for c in r])
    return {"headers":headers,"rows":all_rows,"total_rows":len(all_rows)}

@app.post("/api/admin/upload-excel-save-file")
def upload_excel_save_file(req:Request,file:UploadFile=File(...),tcode_code:str=Form(...),mapping:str=Form(...)):
    u=need_admin(req)
    mapping_dict=json.loads(mapping)
    col_to_field={}
    for k,v in mapping_dict.items():
        if str(k).isdigit(): col_to_field[int(k)]=v
        else:
            try: col_to_field[int(v)]=k
            except: pass
    data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True))
    count=0
    with closing(get_db()) as d:
        for r in rows[1:]:
            if not any(r): continue
            data_dict={}
            for col_idx, field_key in col_to_field.items():
                if col_idx < len(r): data_dict[field_key]=str(r[col_idx] or "").strip()
            if not data_dict: continue
            d.execute("INSERT INTO tcode_records(tcode_code,data,status,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(tcode_code,json.dumps(data_dict),"active",u["username"],u["full_name"],now(),now()))
            count+=1
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

# PLANNING LOOKUP
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
                if str(v).strip().lower()==q or q in str(v).strip().lower():
                    return {"found":True,"data":data,"matched_field":k,"matched_value":v}
                if "job" in k.lower() or "item" in k.lower() or "sap" in k.lower():
                    if q in str(v).lower():
                        return {"found":True,"data":data,"matched_field":k,"matched_value":v}
        rows2=d.execute("SELECT job_no,customer,material_description,specified_length,drum_size FROM rows WHERE job_no LIKE ? LIMIT 1", (f"%{q}%",)).fetchone()
        if rows2:
            return {"found":True,"data":{"job_no":rows2["job_no"],"customer":rows2["customer"],"material_description":rows2["material_description"],"specified_length":rows2["specified_length"],"drum_size":rows2["drum_size"]},"source":"production"}
    return {"found":False}

@app.get("/api/planning/all-jobs")
def all_jobs(req:Request):
    need(req)
    with closing(get_db()) as d:
        rows=[dict(r) for r in d.execute("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active' LIMIT 500").fetchall()]
        jobs=[]
        for r in rows:
            data=json.loads(r["data"])
            job=data.get("job_no") or data.get("item_no") or data.get("sap_no") or data.get("job_item_no") or ""
            if job: jobs.append({"job_no":job,"customer":data.get("customer",""),"material":data.get("material_description") or data.get("material") or ""})
        return jobs[:200]

@app.get("/api/users")
def users(req:Request): need(req); return [dict(r) for r in get_db().execute("SELECT username,full_name,department FROM users WHERE status='approved'").fetchall()]
@app.get("/api/departments")
def depts(req:Request): need(req); return [dict(r) for r in get_db().execute("SELECT * FROM departments").fetchall()]

HTML_PAGE = r"""<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>V10 Full - Planning AutoLink</title><script src="https://cdn.tailwindcss.com"></script><link href="https://cdnjs.cloudflare.com/ajax/libs/remixicon/4.2.0/remixicon.min.css" rel="stylesheet"><style>@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');body{font-family:Inter;background:#f3f4f6}</style></head><body><div id="root"></div><div id="toasts" class="fixed top-3 right-3 z-[9999]"></div><div id="ctxMenu" class="hidden"></div><script>
const $=id=>document.getElementById(id);const toast=(m,t="ok")=>{let e=document.createElement("div");e.className=`px-4 py-3 rounded-xl text-white text-sm font-semibold ${t=="err"?"bg-red-600":"bg-gray-900"}`;e.textContent=m;$("toasts").appendChild(e);setTimeout(()=>e.remove(),3000)};let U=null,ORG=null;async function api(p,o={}){let r=await fetch(p,{credentials:"include",...o});if(!r.ok)throw new Error(await r.text());let ct=r.headers.get("content-type");if(ct&&ct.includes("application/json"))return r.json();return r;}
async function boot(){try{U=await api("/api/me");ORG=await api("/api/org");app()}catch{auth()}}boot();
function auth(){root.innerHTML=`<div class="min-h-screen flex items-center justify-center p-4 bg-[#0f172a]"><div class="bg-white rounded-2xl p-8 w-full max-w-md shadow-2xl"><div class="flex items-center gap-3 mb-2"><div class="w-10 h-10 rounded-xl bg-black text-white grid place-items-center"><i class="ri-shield-keyhole-line"></i></div><div><h2 class="text-2xl font-bold leading-none">ProTransfer V10</h2><div class="text-[11px] text-gray-500">Admin creates users only - No self registration</div></div></div><p class="text-sm text-gray-500 mb-6 mt-4">Login with credentials given by Admin<br><span class="text-xs bg-gray-100 px-2 py-1 rounded">Default: admin / Admin@123</span></p><div id="L"><input id="lu" placeholder="Username" class="w-full border p-3 rounded-xl mb-3 outline-none focus:ring-2 focus:ring-black"><input id="lp" type="password" placeholder="Password" class="w-full border p-3 rounded-xl mb-4 outline-none focus:ring-2 focus:ring-black"><button onclick="login()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">Login</button><div class="mt-4 p-3 bg-gray-50 rounded-xl text-[11px] text-gray-500"><b>Note:</b> No registration. Only Admin can create users from Manage Users panel. Contact Admin for account.</div></div></div></div>`}
async function login(){let f=new FormData();f.append("username",$("lu").value);f.append("password",$("lp").value);try{await api("/api/login",{method:"POST",body:f});location.reload()}catch(e){toast(e.message,"err")}}
function shell(inner){return `<div class="min-h-screen"><div class="bg-white border-b h-[56px] flex items-center justify-center relative"><div class="flex items-center gap-2"><i class="ri-building-4-line"></i><span class="font-bold" id="orgName">${ORG.name}</span>${U.role=='admin'?`<button onclick="editOrg()" class="w-6 h-6 rounded-full bg-gray-100 grid place-items-center"><i class="ri-pencil-line text-xs"></i></button>`:''}</div></div><div class="bg-[#f9fafb] border-b h-[52px] flex items-center justify-between px-6"><div><div class="text-[18px] font-bold">Welcome to dashboard</div><div class="text-[11px] text-gray-500">Layout Code + Title = Sub T-Code • Choose Layout shows all • Planning auto-fills by Job No</div></div><div class="text-right"><div class="text-[14px] font-semibold">${U.full_name}</div><div class="text-[11px] text-gray-500">${U.department} • ${U.role}</div></div></div><div class="bg-white border-b h-[56px] flex items-center justify-between px-4"><div class="flex items-center gap-2"><input id="globalSearch" oninput="filterTcodes(this.value)" placeholder="Search T-Code / Layout Code / Job No..." class="px-4 py-2 rounded-full bg-gray-100 text-sm w-[360px]"></div><div class="flex gap-2 flex-wrap">${U.role=='admin'?`<button onclick="openUserManager()" class="px-3 py-2 rounded-full bg-black text-white text-xs"><i class="ri-user-add-line"></i> Users</button><button onclick="openAccessManager()" class="px-3 py-2 rounded-full bg-indigo-600 text-white text-xs">Access</button><button onclick="openExcelUpload('PLN0')" class="px-3 py-2 rounded-full bg-green-600 text-white text-xs">Excel</button>`:''}<button onclick="v('dash')" class="px-3 py-2 rounded-full bg-gray-900 text-white text-xs">Dashboard</button><button onclick="logout()" class="px-3 py-2 rounded-full bg-white border text-xs">Logout</button></div></div><div class="max-w-[1480px] mx-auto p-5">${inner}</div></div>`}
function app(){v('dash')}
async function v(n){if(n=='dash'){let tcodes=await api("/api/tcodes");$("root").innerHTML=shell(`<div class="flex justify-between mb-4 flex-wrap gap-2"><h2 class="font-bold">T-Codes • ${tcodes.length} • Sub T-Codes = Layouts • Planning Excel Master</h2><span class="text-xs bg-black text-white px-3 py-1 rounded-full">V10 FULL - No Registration, Admin Only</span></div><div id="tcodeGrid" class="grid grid-cols-4 gap-4">${tcodes.map(t=>`<div onclick="openTcode('${t.code}')" class="bg-white rounded-[20px] border p-5 cursor-pointer hover:shadow-xl"><div class="flex justify-between"><div class="w-12 h-12 rounded-xl grid place-items-center text-white" style="background:${t.color}"><i class="${t.icon}"></i></div><span class="text-[10px] px-2 py-1 rounded-full bg-gray-100">${t.layout_count} layouts</span></div><div class="mt-4"><div class="text-[11px] font-bold text-gray-500">${t.code} • ${t.layout_count} sub T-Codes</div><div class="text-[14px] font-bold">${t.title}</div><div class="text-[12px] text-gray-500">${t.description}</div></div></div>`).join('')}</div>`)}}
function filterTcodes(q){document.querySelectorAll("#tcodeGrid > div").forEach(c=>{c.style.display=c.innerText.toLowerCase().includes(q.toLowerCase())?"":"none"})}
async function editOrg(){let nn=prompt("Org Name",ORG.name);if(!nn)return;let f=new FormData();f.append("name",nn);await api("/api/org",{method:"POST",body:f});ORG.name=nn;$("orgName").innerText=nn;toast("Updated")}

async function openTcode(code){
let layouts=await api("/api/tcode/"+code+"/layouts");let tc=(await api("/api/tcodes/all").catch(()=>api("/api/tcodes"))).find(x=>x.code==code)||{code:code,title:code,color:"#000",icon:"ri-apps-line"};
let saved=localStorage.getItem("layout_"+code);let curLayout=layouts.find(l=>l.id==saved)||layouts.find(l=>l.is_default)||layouts[0];
let records=await api("/api/tcode/"+code+"/records");
if(!curLayout){$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6">No layouts for ${code}</div>`);return}
let formHtml=curLayout.fields.map(f=>{
let isJobField=f.key.toLowerCase().includes("job")||f.key.toLowerCase().includes("item")||f.key.toLowerCase().includes("sap");
let onBlur=isJobField?`onblur="lookupPlanning('${f.key}',this.value)"`:"";
let hint=isJobField?`<div id="hint_${f.key}" class="text-[10px] text-green-600 mt-1 hidden">✓ From Planning</div>`:"";
return `<div class="col-span-${f.width||1}"><label class="text-[11px] font-semibold">${f.label}${f.required?' *':''} ${isJobField?'<span class="text-[9px] px-1 py-0.5 rounded bg-blue-50 text-blue-600">🔍 Auto from Planning</span>':''}</label>${f.type=='textarea'?`<textarea id="f_${f.key}" ${onBlur} class="w-full border p-2.5 rounded-xl mt-1" rows="2"></textarea>`:f.type=='select'?`<select id="f_${f.key}" class="w-full border p-2.5 rounded-xl mt-1"><option value="">Select</option>${(f.options||[]).map(o=>`<option>${o}</option>`).join('')}</select>`:f.type=='date'?`<input id="f_${f.key}" type="date" ${onBlur} class="w-full border p-2.5 rounded-xl mt-1">`:f.type=='number'?`<input id="f_${f.key}" type="number" class="w-full border p-2.5 rounded-xl mt-1">`:`<input id="f_${f.key}" type="text" ${onBlur} class="w-full border p-2.5 rounded-xl mt-1" list="jobList">`}${hint}</div>`
}).join('');
formHtml+=`<datalist id="jobList"></datalist>`;
$("root").innerHTML=shell(`
<div class="flex justify-between items-center mb-4 flex-wrap gap-2"><div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl grid place-items-center text-white" style="background:${tc.color}"><i class="${tc.icon}"></i></div><div><div class="text-[11px] font-bold text-gray-500">${tc.code} • ${layouts.length} sub T-Codes (Layouts) • Planning master: PLN0</div><div class="text-[18px] font-bold">${tc.title}</div></div></div><div class="flex gap-2"><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-white border text-sm">← Dashboard</button>${U.role=='admin'?`<button onclick="openExcelUpload('${code}')" class="px-4 py-2 rounded-full bg-green-600 text-white text-sm"><i class="ri-upload-2-line"></i> Upload Excel</button><button onclick="editSchema('${code}')" class="px-4 py-2 rounded-full bg-indigo-600 text-white text-sm">Edit Form</button>`:''}</div></div>
<div class="bg-white rounded-2xl border p-4 mb-4 flex items-center gap-4 flex-wrap"><label class="text-sm font-bold whitespace-nowrap"><i class="ri-layout-line"></i> Choose Layout:</label><select id="layoutSelect" onchange="switchLayout('${code}',this.value)" class="flex-1 border p-2.5 rounded-xl text-sm min-w-[340px]">${layouts.map(l=>`<option value="${l.id}" ${l.id==curLayout.id?'selected':''}>${l.layout_code} - ${l.title} ${l.is_default?' (Default)':''}</option>`).join('')}</select><button onclick="showLayoutDetails()" class="px-3 py-2 rounded-full bg-gray-100 text-sm">Details</button><button onclick="openLayoutMenu(event)" class="px-3 py-2 rounded-full bg-gray-100 text-sm">⋮ Options</button><a href="/api/tcode/${code}/export-excel" class="px-3 py-2 rounded-full bg-gray-100 text-sm"><i class="ri-download-line"></i> Export</a></div>
<div class="grid grid-cols-12 gap-5"><div class="col-span-5 bg-white rounded-2xl border p-5" oncontextmenu="openCtx(event,'${code}')"><h3 class="font-bold text-sm flex justify-between"><span>${curLayout.layout_code} : ${curLayout.title}</span><span class="text-[10px] px-2 py-1 rounded-full ${curLayout.is_default?'bg-green-100 text-green-700':'bg-gray-100'}">${curLayout.is_default?'★ Default':'Custom'}</span></h3><div class="text-[11px] text-gray-400 mt-1">Type Job No and it auto-fills from Planning • Right-click for details</div><div id="formArea" class="mt-4 grid grid-cols-3 gap-3">${formHtml}</div><div class="mt-4 flex gap-2"><button onclick="saveRecord('${code}')" class="flex-1 bg-black text-white py-3 rounded-xl">+ Save</button><button onclick="clearForm()" class="px-5 py-3 rounded-xl bg-gray-100">Clear</button></div><div id="editHint" class="text-xs text-gray-400 mt-2"></div></div><div class="col-span-7 bg-white rounded-2xl border p-5"><div class="flex justify-between"><h3 class="font-bold text-sm">Records • ${records.length} • ${code} • Planning linked by Job No</h3><input id="recSearch" oninput="filterRecords(this.value)" placeholder="Search..." class="border px-3 py-1 rounded-full text-sm"></div><div class="mt-3 overflow-auto max-h-[70vh] border rounded-xl"><table class="w-full text-xs"><thead class="bg-gray-50"><tr><th class="p-2">ID</th><th class="p-2">Layout</th><th class="p-2">Data</th><th class="p-2">By</th><th class="p-2">Action</th></tr></thead><tbody id="recBody">${records.map(r=>`<tr class="border-b"><td class="p-2">#${r.id}</td><td class="p-2 font-mono text-[11px]">${curLayout.layout_code}</td><td class="p-2">${Object.entries(r.data).slice(0,3).map(([k,v])=>`<span class="mr-2"><b>${k}:</b> ${String(v).slice(0,20)}</span>`).join('')}</td><td class="p-2">${r.created_by_name}<div class="text-[10px] text-gray-400">${r.created_at.slice(0,16).replace('T',' ')}</div></td><td class="p-2"><button onclick="loadEdit(${r.id})" class="px-2 py-1 rounded-full bg-gray-100 text-[11px]">Edit</button>${U.role=='admin'?`<button onclick="delRecord('${code}',${r.id})" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Del</button>`:''}</td></tr>`).join('')}</tbody></table></div></div></div>`);
window._curLayouts=layouts;window._curLayout=curLayout;window._curRecords=records;window._curCode=code;
setTimeout(loadJobList,500);
}
function switchLayout(code,id){localStorage.setItem("layout_"+code,id);openTcode(code)}
async function saveAsLayout(code){
let layout_code=prompt("Enter Layout Code (Sub T-Code) e.g. "+code+"-QC02 or "+code+"-MY01\nLeave empty for auto e.g. "+code+"-004","");
let title=prompt("Enter Title for this Sub T-Code e.g. QC Quick Entry","");if(!title){toast("Title required","err");return}
let desc=prompt("Description (optional)","")||"";let fields=window._curLayout.fields;let layout_config=window._curLayout.layout_config||{columns:3};
try{let res=await api("/api/tcode/"+code+"/layouts",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({layout_code:layout_code,name:title,title:title,fields:fields,layout_config:layout_config,description:desc})});toast("Saved Sub T-Code "+res.layout_code+" : "+title);openTcode(code);}catch(e){toast(e.message,"err")}
}
function showLayoutDetails(){
let l=window._curLayout;let html=`<div class="bg-white rounded-2xl border p-5 max-w-md"><h3 class="font-bold">Sub T-Code Details - ${l.layout_code} : ${l.title}</h3><div class="mt-4 space-y-2 text-sm"><div class="flex justify-between"><span class="text-gray-500">Layout Code (Sub T-Code)</span><b class="font-mono">${l.layout_code}</b></div><div class="flex justify-between"><span class="text-gray-500">Title</span><b>${l.title}</b></div><div class="flex justify-between"><span class="text-gray-500">Main T-Code</span><b>${window._curCode}</b></div><div class="flex justify-between"><span class="text-gray-500">Created By</span><b>${l.created_by_name} @${l.created_by}</b></div><div class="flex justify-between"><span class="text-gray-500">Created At</span><b>${new Date(l.created_at).toLocaleString()}</b></div><div class="flex justify-between"><span class="text-gray-500">Last Modified</span><b>${new Date(l.updated_at).toLocaleString()}</b></div><div class="flex justify-between"><span class="text-gray-500">Is Default</span><b>${l.is_default?'Yes ★ Default':'No'}</b></div><div class="flex justify-between"><span class="text-gray-500">Fields</span><b>${l.fields.length}</b></div><div class="mt-3"><span class="text-gray-500 text-xs">Description</span><div class="mt-1 p-2 bg-gray-50 rounded-xl">${l.description||'No description'}</div></div><div class="mt-3"><span class="text-gray-500 text-xs">Fields List</span><div class="mt-1 text-[11px] bg-gray-50 p-2 rounded-xl">${l.fields.map(f=>f.key+':'+f.label).join(', ')}</div></div></div><button onclick="hideCtx()" class="mt-4 w-full bg-black text-white py-2 rounded-xl">Close</button></div>`;
let ctx=$("ctxMenu");ctx.innerHTML=html;ctx.className="fixed inset-0 bg-black/40 grid place-items-center z-[9999]";ctx.classList.remove("hidden");
}
function openLayoutMenu(e){
e.stopPropagation();let l=window._curLayout;let code=window._curCode;
let html=`<div style="position:fixed;left:${e.clientX}px;top:${e.clientY}px;background:white;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 20px 40px rgba(0,0,0,.15);padding:6px;min-width:280px"><div class="p-2 space-y-1"><div class="px-3 py-2 font-bold text-sm">${l.layout_code} : ${l.title}</div><button onclick="showLayoutDetails();hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm"><i class="ri-information-line"></i> Details - Who, When, Last Modified</button><button onclick="saveAsLayout('${code}');hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm"><i class="ri-save-line"></i> Save as New Sub T-Code</button><button onclick="setDefaultLayout('${code}');hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm"><i class="ri-star-line"></i> Set as Default</button><button onclick="deleteLayout('${code}',${l.id});hideCtx()" class="w-full text-left px-3 py-2 hover:bg-red-50 text-red-600 rounded-xl text-sm"><i class="ri-delete-bin-line"></i> Delete ${l.layout_code}</button></div></div>`;
let ctx=$("ctxMenu");ctx.innerHTML=html;ctx.className="fixed inset-0 z-[9999]";ctx.classList.remove("hidden");
}
function openCtx(e,code){
e.preventDefault();let l=window._curLayout;
let html=`<div style="position:fixed;left:${e.clientX}px;top:${e.clientY}px;background:white;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 20px 40px rgba(0,0,0,.15);padding:8px;min-width:300px"><div class="px-3 py-2 font-bold text-sm">${l.layout_code} : ${l.title} ${l.is_default?'★':''}</div><div class="px-3 py-1 text-[11px] text-gray-500">Created By: ${l.created_by_name} @${l.created_by}</div><div class="px-3 py-1 text-[11px] text-gray-500">Created: ${new Date(l.created_at).toLocaleString()}</div><div class="px-3 py-1 text-[11px] text-gray-500">Modified: ${new Date(l.updated_at).toLocaleString()}</div><div class="px-3 py-1 text-[11px] text-gray-500">Fields: ${l.fields.length} • ${l.description||''}</div><div class="border-t my-2"></div><button onclick="showLayoutDetails();hideCtx()" class="w-full text-left px-3 py-2 hover:bg-gray-100 rounded-xl text-sm">Full Details</button></div>`;
let ctx=$("ctxMenu");ctx.innerHTML=html;ctx.className="fixed inset-0 z-[9999]";ctx.classList.remove("hidden");
}
function hideCtx(){$("ctxMenu").classList.add("hidden");$("ctxMenu").innerHTML=""}
document.addEventListener("click",()=>{if(!$("ctxMenu").classList.contains("hidden"))hideCtx()})
async function setDefaultLayout(code){if(U.role!='admin'){toast("Admin only","err");return}let id=window._curLayout.id;await api("/api/tcode/"+code+"/layouts/"+id+"/set-default",{method:"POST"});toast("Set default "+window._curLayout.layout_code);openTcode(code)}
async function deleteLayout(code,id){if(window._curLayout.is_default){toast("Cannot delete default","err");return}if(!confirm("Delete "+window._curLayout.layout_code+" : "+window._curLayout.title+"?"))return;await api("/api/tcode/"+code+"/layouts/"+id,{method:"DELETE"});localStorage.removeItem("layout_"+code);toast("Deleted");openTcode(code)}
let _editId=null;
function clearForm(){_editId=null;document.querySelectorAll("#formArea input, #formArea textarea, #formArea select").forEach(e=>e.value="");$("editHint").innerText=""}
function loadEdit(id){let rec=window._curRecords.find(x=>x.id==id);if(!rec)return;_editId=id;for(let k in rec.data){let el=document.getElementById("f_"+k);if(el)el.value=rec.data[k]}$("editHint").innerText="Editing #"+id;window.scrollTo(0,0)}
async function saveRecord(code){let data={};for(let f of window._curLayout.fields){let el=document.getElementById("f_"+f.key);if(!el)continue;if(f.required&&!el.value){toast(f.label+" required","err");return}data[f.key]=el.value}try{if(_editId){await api("/api/tcode/"+code+"/records/"+_editId,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});toast("Updated")}else{await api("/api/tcode/"+code+"/records",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});toast("Saved")}clearForm();openTcode(code)}catch(e){toast(e.message,"err")}}
async function delRecord(code,id){if(!confirm("Delete #"+id+"?"))return;await api("/api/tcode/"+code+"/records/"+id,{method:"DELETE"});toast("Deleted");openTcode(code)}
function filterRecords(q){document.querySelectorAll("#recBody tr").forEach(r=>{r.style.display=r.innerText.toLowerCase().includes(q.toLowerCase())?"":"none"})}

async function editSchema(code){let schema=await api("/api/tcode/"+code+"/schema");let fields=schema.fields;let listHtml=fields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('');$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold">Edit Form - ${code} → Updates ${code}-001</h2><div class="flex gap-2"><button onclick="openTcode('${code}')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button><button onclick="saveSchema('${code}')" class="px-4 py-2 rounded-full bg-black text-white text-sm">Save</button></div></div><div class="mt-6 grid grid-cols-12 gap-5"><div class="col-span-7"><div id="fieldList" class="space-y-2">${listHtml}</div></div><div class="col-span-5 bg-gray-50 rounded-2xl p-4 border"><h3 class="font-semibold text-sm">Add Field</h3><div class="mt-3 space-y-3"><input id="nf_key" placeholder="key" class="w-full border p-2.5 rounded-xl text-sm"><input id="nf_label" placeholder="Label" class="w-full border p-2.5 rounded-xl text-sm"><select id="nf_type" class="w-full border p-2.5 rounded-xl text-sm"><option value="text">Text</option><option value="number">Number</option><option value="date">Date</option><option value="textarea">Long Text</option><option value="select">Dropdown</option></select><input id="nf_options" placeholder="Options comma" class="w-full border p-2.5 rounded-xl text-sm"><div class="grid grid-cols-2 gap-3"><select id="nf_width" class="border p-2.5 rounded-xl text-sm"><option value="1">W1</option><option value="2">W2</option><option value="3">W Full</option></select><label class="flex items-center gap-2 text-sm"><input id="nf_req" type="checkbox"> Required</label></div><button onclick="addField()" class="w-full bg-black text-white py-2.5 rounded-xl text-sm">+ Add</button></div></div></div></div>`);window._editFields=fields}
function addField(){let key=$("nf_key").value.trim().replace(/\s+/g,'_').toLowerCase();if(!key){toast("Enter key","err");return}let field={key:key,label:$("nf_label").value||key,type:$("nf_type").value,width:parseInt($("nf_width").value),required:$("nf_req").checked};if(field.type=='select'){field.options=$("nf_options").value.split(',').map(s=>s.trim()).filter(Boolean)}window._editFields.push(field);$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('')}
function removeField(i){window._editFields.splice(i,1);$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('')}
function moveField(i,dir){let j=i+dir;if(j<0||j>=window._editFields.length)return;let t=window._editFields[i];window._editFields[i]=window._editFields[j];window._editFields[j]=t;$("fieldList").innerHTML=window._editFields.map((f,i)=>`<div class="flex items-center gap-2 p-2 bg-gray-50 rounded-xl text-xs"><span>${i+1}</span><b>${f.key}</b> ${f.label}<div class="ml-auto flex gap-1"><button onclick="moveField(${i},-1)" class="px-2 py-1 bg-white border rounded">↑</button><button onclick="moveField(${i},1)" class="px-2 py-1 bg-white border rounded">↓</button><button onclick="removeField(${i})" class="px-2 py-1 bg-red-50 text-red-600 border rounded">X</button></div></div>`).join('')}
async function saveSchema(code){await api("/api/tcode/"+code+"/schema",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({fields:window._editFields,layout:{columns:3}})});toast("Saved");openTcode(code)}

// ADMIN USER MANAGER
async function openUserManager(){
let users=await api("/api/admin/users");let depts=await api("/api/departments");
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between items-center"><h2 class="font-bold text-lg"><i class="ri-user-settings-line"></i> Manage Users - Only Admin Creates Users (No Registration)</h2><div class="flex gap-2"><button onclick="openAccessManager()" class="px-4 py-2 rounded-full bg-indigo-600 text-white text-sm">Manage Access</button><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div></div>
<div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-5 bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Create New User (Admin Only)</h3><div class="mt-4 space-y-3">
<input id="new_username" placeholder="Username e.g. ramesh" class="w-full border p-2.5 rounded-xl text-sm"><input id="new_password" placeholder="Password e.g. Ramesh@123" class="w-full border p-2.5 rounded-xl text-sm"><input id="new_fullname" placeholder="Full Name e.g. Ramesh Kumar" class="w-full border p-2.5 rounded-xl text-sm">
<select id="new_dept" class="w-full border p-2.5 rounded-xl text-sm"><option value="">Select Department</option>${depts.map(d=>`<option value="${d.name}">${d.name}</option>`).join('')}</select>
<input id="new_plant" placeholder="Plant e.g. Gooty" value="Gooty" class="w-full border p-2.5 rounded-xl text-sm">
<select id="new_role" class="w-full border p-2.5 rounded-xl text-sm"><option value="user">User</option><option value="admin">Admin</option></select>
<button onclick="doCreateUser()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">+ Create User</button>
<div class="text-[11px] text-gray-500 mt-2">• User will be able to login immediately<br>• No approval needed<br>• Then grant T-Code access from Manage Access</div>
</div></div>
<div class="col-span-7"><h3 class="font-bold text-sm">All Users • ${users.length} • No self-registration</h3><div class="mt-3 border rounded-xl overflow-auto max-h-[70vh]"><table class="w-full text-xs"><thead class="bg-gray-50 sticky top-0"><tr><th class="p-2 text-left">Username</th><th class="p-2">Name</th><th class="p-2">Dept</th><th class="p-2">Role</th><th class="p-2">Action</th></tr></thead><tbody>${users.map(u=>`<tr class="border-b hover:bg-gray-50"><td class="p-2 font-mono font-bold">@${u.username}<div class="text-[10px] text-gray-400">${u.plant}</div></td><td class="p-2">${u.full_name}</td><td class="p-2">${u.department}</td><td class="p-2"><span class="px-2 py-1 rounded-full ${u.role=='admin'?'bg-black text-white':'bg-gray-100'} text-[10px]">${u.role}</span></td><td class="p-2 flex gap-1"><button onclick="resetUserPass('${u.username}')" class="px-2 py-1 rounded-full bg-gray-100 text-[11px]">Reset Pass</button><button onclick="deleteUser('${u.username}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Delete</button></td></tr>`).join('')}</tbody></table></div></div></div></div>`);
}
async function doCreateUser(){
let f=new FormData();f.append("username",$("new_username").value);f.append("password",$("new_password").value);f.append("full_name",$("new_fullname").value);f.append("department",$("new_dept").value);f.append("plant",$("new_plant").value);f.append("role",$("new_role").value);
if(!f.get("username")||!f.get("password")){toast("Username and password required","err");return}
try{await api("/api/admin/create-user",{method:"POST",body:f});toast("User created: @"+f.get("username"));openUserManager()}catch(e){toast(e.message,"err")}
}
async function deleteUser(username){if(!confirm("Delete user @"+username+"? This will also delete his access grants."))return;let f=new FormData();f.append("username",username);await api("/api/admin/delete-user",{method:"POST",body:f});toast("Deleted");openUserManager()}
async function resetUserPass(username){let np=prompt("Enter new password for @"+username);if(!np)return;let f=new FormData();f.append("username",username);f.append("new_password",np);await api("/api/admin/reset-password",{method:"POST",body:f});toast("Password reset for @"+username)}

// ACCESS MANAGER
async function openAccessManager(){
let users=await api("/api/admin/users");let tcodes=await api("/api/tcodes/all");let grants=await api("/api/admin/access-grants");
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg"><i class="ri-shield-keyhole-line"></i> Manage Access - Give T-Code / Planning / Layout to Any User (Admin)</h2><div class="flex gap-2"><button onclick="openUserManager()" class="px-4 py-2 rounded-full bg-black text-white text-sm">Users</button><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div></div>
<div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-5 bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Grant Access</h3><div class="mt-4 space-y-3">
<select id="acc_user" class="w-full border p-2.5 rounded-xl text-sm"><option value="">-- Select User --</option>${users.map(u=>`<option value="${u.username}">${u.full_name} @${u.username} - ${u.department}</option>`).join('')}</select>
<select id="acc_tcode" class="w-full border p-2.5 rounded-xl text-sm" onchange="loadLayoutsForAccess(this.value)"><option value="">-- Select T-Code (e.g. PLN0) --</option>${tcodes.map(t=>`<option value="${t.code}">${t.code} - ${t.title}</option>`).join('')}</select>
<select id="acc_layout" class="w-full border p-2.5 rounded-xl text-sm"><option value="">-- All Layouts of this T-Code (Wildcard) --</option></select>
<div class="text-[11px] text-gray-500">• If Layout empty = user gets ALL layouts (Sub T-Codes) of that T-Code<br>• If you select specific Layout Code = user gets only that Sub T-Code e.g. PLN0-001</div>
<button onclick="doGrantAccess()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">+ Grant Access</button>
</div></div>
<div class="col-span-7"><h3 class="font-bold text-sm">Current Grants • ${grants.length}</h3><div class="mt-3 border rounded-xl overflow-auto max-h-[70vh]"><table class="w-full text-xs"><thead class="bg-gray-50 sticky top-0"><tr><th class="p-2 text-left">User</th><th class="p-2">T-Code</th><th class="p-2">Layout (Sub T-Code)</th><th class="p-2">Granted By</th><th class="p-2">Action</th></tr></thead><tbody>${grants.map(g=>`<tr class="border-b"><td class="p-2"><b>${g.full_name||g.username}</b><div class="text-[10px] text-gray-500">@${g.username} • ${g.department||''}</div></td><td class="p-2 font-mono font-bold">${g.tcode_code}</td><td class="p-2 font-mono">${g.layout_code||'<span class="text-gray-400">ALL Layouts</span>'}</td><td class="p-2">${g.granted_by}<div class="text-[10px] text-gray-400">${g.granted_at.slice(0,16).replace('T',' ')}</div></td><td class="p-2"><button onclick="revokeGrant(${g.id})" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Revoke</button></td></tr>`).join('')}</tbody></table></div><div class="mt-3 text-[11px] text-gray-500">Example: Give PLN0 to store user → He will see Planning Details (including your Excel uploaded data) on his dashboard.</div></div></div></div>`);
}
async function loadLayoutsForAccess(tcode){
let sel=$("acc_layout");sel.innerHTML='<option value="">-- All Layouts --</option>';if(!tcode)return;
let layouts=await api("/api/tcode/"+tcode+"/layouts");sel.innerHTML='<option value="">-- All Layouts of this T-Code (Wildcard) --</option>'+layouts.map(l=>`<option value="${l.layout_code}">${l.layout_code} - ${l.title} ${l.is_default?' (Default)':''}</option>`).join('');
}
async function doGrantAccess(){
let user=$("acc_user").value;let tcode=$("acc_tcode").value;let layout=$("acc_layout").value;
if(!user||!tcode){toast("Select User and T-Code","err");return}
let f=new FormData();f.append("username",user);f.append("tcode_code",tcode);f.append("layout_code",layout);
try{await api("/api/admin/grant-access",{method:"POST",body:f});toast(`Access granted: ${user} → ${tcode} ${layout||'ALL'}`);openAccessManager()}catch(e){toast(e.message,"err")}
}
async function revokeGrant(id){if(!confirm("Revoke this access?"))return;let f=new FormData();f.append("grant_id",id);await api("/api/admin/revoke-access",{method:"POST",body:f});toast("Revoked");openAccessManager()}

// EXCEL UPLOAD
async function openExcelUpload(code){
$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6 max-w-4xl mx-auto"><div class="flex justify-between"><h2 class="font-bold text-lg">Upload Excel to ${code} - Planning Details (Admin)</h2><button onclick="openTcode('${code}')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div>
<div class="mt-6 space-y-4"><div class="border-2 border-dashed rounded-2xl p-8 text-center bg-gray-50"><input id="excelFile" type="file" accept=".xlsx,.xls" class="hidden" onchange="previewExcel('${code}')"><button onclick="document.getElementById('excelFile').click()" class="px-6 py-3 bg-black text-white rounded-xl"><i class="ri-upload-2-line"></i> Choose Excel File</button><div class="text-xs text-gray-500 mt-2">First row should be headers like Job No, Customer, Material, Length, etc.</div><div id="fileName" class="mt-2 text-sm font-semibold"></div></div><div id="excelPreview" class="hidden"><h3 class="font-bold text-sm">Preview - Map Excel Columns to Form Fields</h3><div class="mt-3 overflow-auto border rounded-xl max-h-[40vh]"><table class="w-full text-xs" id="previewTable"></table></div><div id="mappingArea" class="mt-4 grid grid-cols-3 gap-3"></div><button onclick="saveExcelToPlanning('${code}')" class="mt-4 w-full bg-green-600 text-white py-3 rounded-xl font-semibold">Load ${code} - All Rows to Planning Details</button></div></div></div>`);
}
let _excelData={};let _uploadFile=null;
async function previewExcel(code){
let fileEl=$("excelFile");if(!fileEl.files[0])return;$("fileName").innerText=fileEl.files[0].name;
let f=new FormData();f.append("file",fileEl.files[0]);let res=await api("/api/admin/upload-excel-preview",{method:"POST",body:f});
let fields=(await api("/api/tcode/"+code+"/schema")).fields;
let mappingHtml=fields.map(ff=>`<div><label class="text-[11px] font-semibold">${ff.label} (${ff.key})</label><select id="map_${ff.key}" class="w-full border p-2 rounded-xl text-xs mt-1"><option value="">-- Ignore --</option>${res.headers.map((h,i)=>`<option value="${i}" ${h.toLowerCase().includes(ff.key.toLowerCase())||h.toLowerCase().includes(ff.label.toLowerCase().split(' ')[0])?'selected':''}>${h} (Col ${i+1})</option>`).join('')}</select></div>`).join('');
$("mappingArea").innerHTML=mappingHtml;
let tableHtml=`<thead class="bg-gray-900 text-white"><tr>${res.headers.map(h=>`<th class="p-2 text-left">${h}</th>`).join('')}</tr></thead><tbody>${res.preview.map(r=>`<tr class="border-b">${r.map(c=>`<td class="p-2">${c}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td colspan="${res.headers.length}" class="p-2 text-center text-gray-500">Showing 10 of ${res.total_rows} rows - All ${res.total_rows} will be loaded</td></tr></tfoot>`;
$("previewTable").innerHTML=tableHtml;$("excelPreview").classList.remove("hidden");window._uploadFile=fileEl.files[0];
}
async function saveExcelToPlanning(code){
if(!window._uploadFile){toast("Choose file first","err");return}
let fields=(await api("/api/tcode/"+code+"/schema")).fields;let mapping={};
fields.forEach(ff=>{let sel=$("map_"+ff.key);if(sel&&sel.value!==""){mapping[sel.value]=ff.key}});
let fd2=new FormData();fd2.append("file",window._uploadFile);fd2.append("tcode_code",code);fd2.append("mapping",JSON.stringify(mapping));
let saveRes=await api("/api/admin/upload-excel-save-file",{method:"POST",body:fd2});
toast(`Loaded ${saveRes.saved} rows to ${code}`);openTcode(code);
}

// PLANNING LOOKUP
async function lookupPlanning(fieldKey,value){
if(!value||value.length<2)return;
try{
let res=await api("/api/planning/lookup?job_no="+encodeURIComponent(value.trim()));
if(!res.found){res=await api("/api/planning/lookup?item_no="+encodeURIComponent(value.trim()));}
if(!res.found){res=await api("/api/planning/lookup?sap_no="+encodeURIComponent(value.trim()));}
if(res.found){
let data=res.data;
for(let k in data){
let el=document.getElementById("f_"+k);
if(!el){
 for(let fk of window._curLayout.fields){
  if(fk.key.toLowerCase().includes(k.toLowerCase())||k.toLowerCase().includes(fk.key.toLowerCase())){
   el=document.getElementById("f_"+fk.key);if(el)break;
  }
  if(k.toLowerCase().includes("customer")&&fk.key.toLowerCase().includes("customer")){el=document.getElementById("f_"+fk.key)}
  if((k.toLowerCase().includes("material")||k.toLowerCase().includes("desc"))&&fk.key.toLowerCase().includes("material")){el=document.getElementById("f_"+fk.key)}
  if(k.toLowerCase().includes("length")&&fk.key.toLowerCase().includes("length")){el=document.getElementById("f_"+fk.key)}
  if(k.toLowerCase().includes("drum")&&fk.key.toLowerCase().includes("drum")){el=document.getElementById("f_"+fk.key)}
 }
}
if(el&&!el.value&&data[k]){
el.value=data[k];el.classList.add("bg-green-50","border-green-300");setTimeout(()=>{el.classList.remove("bg-green-50","border-green-300")},2000);
}
}
let hintEl=$("hint_"+fieldKey);if(hintEl){hintEl.innerText=`✓ From Planning: ${res.matched_field}=${res.matched_value}`;hintEl.classList.remove("hidden")}
toast("Auto-filled from Planning: Job "+value);
}
}catch(e){console.log(e)}
}
async function loadJobList(){
try{let jobs=await api("/api/planning/all-jobs");let dl=$("jobList");if(!dl)return;dl.innerHTML=jobs.map(j=>`<option value="${j.job_no}">${j.customer} - ${j.material.slice(0,30)}</option>`).join('');}catch(e){}
}

async function logout(){await api("/api/logout",{method:"POST"});location.reload()}
</script></body></html>
"""

@app.get("/",response_class=HTMLResponse)
def index(): return HTML_PAGE

if __name__=="__main__":
    import uvicorn
    port=int(os.environ.get("PORT",8000))
    print(f"\n✅ V10 FULL running at http://localhost:{port}\nDefault admin: admin / Admin@123")
    uvicorn.run(app,host="0.0.0.0",port=port)