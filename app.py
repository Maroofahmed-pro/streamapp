"""
Backend proxy for sindhstream.org.pk
- User apna sindhstream login yahan deta hai (username/password)
- Backend unki taraf se sindhstream pe login karta hai, cookie save karta hai
- Phir same cookie se data fetch karke frontend ko deta hai

Run: pip install flask requests beautifulsoup4
     python app.py
"""
from flask import Flask, request, jsonify, session, render_template_string
import requests
from bs4 import BeautifulSoup

app = Flask(__name__)
app.secret_key = "change-this-to-something-random"  # session cookie encrypt karne ke liye

PAGE = """
<!DOCTYPE html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>STEDA Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.4/jspdf.plugin.autotable.min.js"></script>
<style>
:root{--bg:#f1f4f9;--card:#fff;--bd:#e5e8ee;--tx:#1a1d21;--mu:#6b7280;--ac:#2563eb;--ac2:#1d4ed8;--ok:#16a34a;--ok2:#15803d}
*{box-sizing:border-box}
body{font:14px/1.5 -apple-system,system-ui,Segoe UI,Roboto,sans-serif;background:var(--bg);margin:0;padding:18px;color:var(--tx)}
.card{background:var(--card);border:1px solid var(--bd);border-radius:14px;padding:18px;margin-bottom:14px;box-shadow:0 1px 3px rgba(16,24,40,.04)}
input,select,button{padding:9px 11px;border:1px solid #d5d9e0;border-radius:8px;font:inherit;color:var(--tx);background:#fff}
input:focus,select:focus{outline:none;border-color:var(--ac);box-shadow:0 0 0 3px rgba(37,99,235,.12)}
button{background:var(--ac);color:#fff;border:none;cursor:pointer;font-weight:600;transition:.15s}
button:hover{background:var(--ac2)}
button.ex{background:var(--ok)}
button.ex:hover{background:var(--ok2)}
button.ghost{background:#fff;color:var(--ac);border:1px solid #d5d9e0;font-weight:500}
button.ghost:hover{background:#f3f6fb}
button:disabled{opacity:.6;cursor:default}
#err{color:#dc2626;font-size:13px;margin-top:8px;min-height:16px}
.filters{display:grid;grid-template-columns:2fr 1fr 1fr;gap:10px}
.filters2{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:10px}
@media(max-width:700px){.filters,.filters2{grid-template-columns:minmax(0,1fr)}}
.topbar{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}
.topbar h3{margin:0;font-size:17px}
.info{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.export-actions{display:flex;gap:8px}
.export-error{display:none;color:#dc2626;font-size:13px;margin:0 0 10px}
.count-badge{display:inline-block;background:linear-gradient(135deg,#2563eb,#1d4ed8);color:#fff;font-weight:700;font-size:14px;padding:7px 16px;border-radius:99px;box-shadow:0 2px 6px rgba(37,99,235,.3)}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;background:#f8fafc;color:#475569;padding:10px;white-space:nowrap;font-weight:600;border-bottom:2px solid var(--bd);position:sticky;top:0}
td{padding:10px;border-top:1px solid #f0f1f4;white-space:nowrap}
tbody tr:hover{background:#f8fafc}
.tblwrap{width:100%;max-width:100%;overflow:auto;border:1px solid var(--bd);border-radius:10px;max-height:60vh}
table{min-width:900px}
.pager{display:flex;justify-content:space-between;align-items:center;margin-top:14px}
#pg{color:var(--mu);font-size:13px;font-weight:500}
#dash{display:none}
.b{display:inline-block;font-size:11px;font-weight:600;padding:3px 9px;border-radius:99px;background:#eef2ff;color:#3730a3}
.b.sub{background:#dcfce7;color:#166534}
.login-wrap{min-height:80vh;display:flex;align-items:center;justify-content:center}
.login-card{width:100%;max-width:380px;text-align:center}
.login-card .icon{font-size:36px;margin-bottom:6px}
.login-card h3{margin:0 0 4px;font-size:20px}
.login-card p{color:var(--mu);margin:0 0 20px;font-size:13px}
.login-card input{width:100%;margin-bottom:12px}
.login-card button{width:100%;padding:11px}
input,select{min-width:0;width:100%}
@media(max-width:600px){
  body{padding:10px}
  .card{padding:14px;border-radius:12px}
  .topbar{align-items:flex-start;gap:10px}
  .topbar h3{font-size:16px}
  .topbar button{flex-shrink:0}
  .info{align-items:stretch}
  .count-badge{text-align:center}
  .info .export-actions{width:100%;flex-direction:column}
  .info .ex{width:100%;min-height:44px}
  .filters,.filters2{gap:8px}
  .filters2{margin-top:8px}
  .filters input,.filters select,.filters2 select{min-height:44px}
  .pager{display:grid;grid-template-columns:1fr 1fr;gap:10px}
  #pg{grid-column:1/-1;grid-row:1;text-align:center}
  .pager button{min-height:44px}
  .login-wrap{min-height:calc(100vh - 20px)}
}
</style></head><body>

<div class="login-wrap" id="loginWrap">
  <div class="card login-card" id="loginBox">
    <div class="icon">&#128272;</div>
    <h3>STEDA Dashboard</h3>
    <p>Apne Sindh STREAM account se login karein</p>
    <input id="u" placeholder="Username">
    <input id="p" type="password" placeholder="Password">
    <button onclick="login()">Login</button>
    <div id="err"></div>
  </div>
</div>

<div id="dash">
  <div class="card">
    <div class="topbar"><h3>Project Registration Details</h3>
      <button class="ghost" onclick="loadData(1)">&#8635; Refresh</button></div>
    <div class="filters">
      <input id="q" placeholder="Search project ID, title, school, teacher or SEMIS">
      <select id="entity"><option value="">All entities</option>
        <option>SE&amp;LD</option><option>CED</option><option>SEF</option>
        <option>Private</option><option>NFE</option><option>Madarsah</option></select>
      <select id="district"><option value="">All Districts</option><option value="KMR">Kashmore</option></select>
    </div>
    <div class="filters2">
      <select id="level"><option value="">All levels</option><option>Primary</option><option>Middle</option><option>High</option><option>Higher Secondary</option></select>
      <select id="theme"><option value="">All themes</option>
        <option value="1">1 - Climate Resilience, Environment, Alternative Energy and Green Solutions</option>
        <option value="2">2 - AI, Digital Solution, Gamification, Automation, Coding and Cyber Security</option>
        <option value="3">3 - Technology, Robotics and Internet of Things</option>
        <option value="4">4 - Art, Music, Culture, Heritage and Literature</option>
        <option value="5">5 - Natural Science and Health Science</option></select>
      <select id="status"><option value="">All statuses</option><option value="Submitted">Submitted</option><option value="Pending">Pending / Draft</option></select>
    </div>
    <button onclick="loadData(1)" style="width:100%;margin-top:12px">&#8635; Apply Filters</button>
  </div>

  <div class="card">
    <div class="info"><span class="count-badge" id="cnt">0 entries found</span>
      <div class="export-actions">
        <button class="ex" id="exBtn" onclick="exportExcel()">&#8595; Export Excel</button>
        <button id="pdfBtn" onclick="exportPDF()">&#8595; Export PDF</button>
      </div></div>
    <div class="export-error" id="exportErr" role="alert"></div>
    <div class="tblwrap"><table>
      <thead><tr><th>Project ID</th><th>Project Title</th><th>Entity</th><th>Level</th><th>Theme</th><th>School / Institution</th><th>Teacher</th><th>Tehsil</th><th>Status</th></tr></thead>
      <tbody id="tbody"></tbody>
    </table></div>
    <div class="pager"><button class="ghost" onclick="prevPage()">&larr; Prev</button><span id="pg"></span><button class="ghost" onclick="nextPage()">Next &rarr;</button></div>
  </div>
</div>

<script>
let page=1, totalPages=1, totalRows=0;
async function login(){
  const r = await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({username:document.getElementById('u').value, password:document.getElementById('p').value})});
  const d = await r.json();
  if(d.success){document.getElementById('loginWrap').style.display='none';document.getElementById('dash').style.display='block';loadData(1);}
  else document.getElementById('err').textContent = d.message || 'Login failed';
}
function filterParams(page,pageSize){
  return new URLSearchParams({
    search: document.getElementById('q').value,
    entity: document.getElementById('entity').value,
    district: document.getElementById('district').value,
    level: document.getElementById('level').value,
    themeNo: document.getElementById('theme').value,
    status: document.getElementById('status').value,
    page, pageSize
  });
}
async function loadData(p){
  page=p;
  const r = await fetch('/api/projects?'+filterParams(page,25));
  if(r.status===401){document.getElementById('dash').style.display='none';document.getElementById('loginWrap').style.display='flex';
    document.getElementById('err').textContent='Session expired, dobara login karo';return;}
  const d = await r.json();
  totalPages = d.TotalPages || 1; totalRows = d.TotalRows || 0;
  document.getElementById('cnt').textContent = totalRows + ' entries found';
  document.getElementById('pg').textContent = `Page ${d.Page}/${d.TotalPages}`;
  document.getElementById('tbody').innerHTML = (d.Projects||[]).map(x=>`<tr>
    <td><b>${x.ProjectID||''}</b></td><td>${x.ProjectTitle||''}</td><td>${x.Entity||''}</td>
    <td>${x.Level||''}</td><td>${x.ThemeNo||''} - ${x.Theme||''}</td><td>${x.SchoolName||''}</td>
    <td>${x.TeacherName||''}</td><td>${x.Tehsil||''}</td>
    <td><span class="b">${x.isSubmitted? 'Submitted':'Not submitted'}</span></td></tr>`).join('')
    || '<tr><td colspan="9">No results</td></tr>';
}
function prevPage(){ if(page>1) loadData(page-1); }
function nextPage(){ if(page<totalPages) loadData(page+1); }

const exportColumns = [
  { header:'Project ID', value:x=>x.ProjectID||'' },
  { header:'Project Title', value:x=>x.ProjectTitle||'' },
  { header:'Entity', value:x=>x.Entity||'' },
  { header:'Level', value:x=>x.Level||'' },
  { header:'Theme', value:x=>`${x.ThemeNo||''} - ${x.Theme||''}` },
  { header:'School / Institution', value:x=>x.SchoolName||'' },
  { header:'Teacher', value:x=>x.TeacherName||'' },
  { header:'Tehsil', value:x=>x.Tehsil||'' },
  { header:'Status', value:x=>x.isSubmitted?'Submitted':'Not submitted' }
];

async function getAllFilteredProjects(btn){
  let all=[], p=1, tp=1;
  do{
    btn.textContent = `Fetching page ${p}${tp>1?'/'+tp:''}...`;
    const r = await fetch('/api/projects?'+filterParams(p,100));
    if(!r.ok) throw new Error(`Could not fetch projects (HTTP ${r.status}).`);
    const d = await r.json();
    tp = d.TotalPages || 1;
    all = all.concat(d.Projects||[]);
    p++;
  } while(p<=tp);
  return all;
}

async function exportData(buttonId, defaultLabel, createFile){
  const btn=document.getElementById(buttonId);
  const error=document.getElementById('exportErr');
  btn.disabled=true;
  error.style.display='none';
  error.textContent='';
  try{
    const all=await getAllFilteredProjects(btn);
    createFile(all);
  } catch(e) {
    error.textContent=e.message || 'Export failed. Please try again.';
    error.style.display='block';
  } finally {
    btn.disabled=false;
    btn.textContent=defaultLabel;
  }
}

function exportExcel(){
  return exportData('exBtn','\u2193 Export Excel',all=>{
    const headers=exportColumns.map(column=>column.header);
    const rows=all.map(project=>exportColumns.map(column=>column.value(project)));
    const ws=XLSX.utils.aoa_to_sheet([headers,...rows]);
    ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:rows.length,c:headers.length-1}})};
    ws['!cols']=[16,30,16,16,42,30,24,18,18].map(wch=>({wch}));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Projects');
    XLSX.writeFile(wb,'steda_projects.xlsx');
  });
}

function exportPDF(){
  return exportData('pdfBtn','\u2193 Export PDF',all=>{
    if(!window.jspdf || !window.jspdf.jsPDF) throw new Error('PDF export library did not load. Please refresh and try again.');
    const {jsPDF}=window.jspdf;
    const pdfColumns=exportColumns.filter(column=>column.header!=='Teacher');
    const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
    doc.setFontSize(14);
    doc.text('Project Registration Details',9,12);
    doc.autoTable({
      startY:17,
      head:[pdfColumns.map(column=>column.header)],
      body:all.map(project=>pdfColumns.map(column=>column.value(project))),
      margin:{left:9,right:9},
      styles:{fontSize:7,cellPadding:2,overflow:'linebreak'},
      headStyles:{fillColor:[37,99,235]},
      columnStyles:{
        0:{cellWidth:22},1:{cellWidth:38},2:{cellWidth:22},3:{cellWidth:18},
        4:{cellWidth:48},5:{cellWidth:52},6:{cellWidth:25},7:{cellWidth:32}
      },
      horizontalPageBreak:true,
      horizontalPageBreakRepeat:0
    });
    doc.save('steda_projects.pdf');
  });
}
</script>
</body></html>
"""


@app.route("/")
def index():
    return render_template_string(PAGE)

BASE = "https://sindhstream.org.pk"
LOGIN_PAGE = f"{BASE}/Login.aspx?ReturnUrl=%2fcontrol%2fdashboard.aspx"
LOGIN_ACTION = LOGIN_PAGE
DASHBOARD_API = f"{BASE}/handlers/Handler_STREAM_Project_Details.ashx"

# Har logged-in user ka cookie yahan rakhte hain (demo ke liye memory mein;
# production mein database/redis use karo)
USER_SESSIONS = {}


@app.route("/api/login", methods=["POST"])
def login():
    data = request.json
    username = data.get("username")
    password = data.get("password")

    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0"})

    # Step 1: login page load karo, ASP.NET ke hidden tokens nikalo
    r1 = s.get(LOGIN_PAGE)
    soup = BeautifulSoup(r1.text, "html.parser")

    def get_val(name):
        tag = soup.find("input", {"name": name})
        return tag["value"] if tag else ""

    payload = {
        "__EVENTTARGET": "HeadLoginView$logId$LoginButton",
        "__EVENTARGUMENT": "",
        "__VIEWSTATE": get_val("__VIEWSTATE"),
        "__VIEWSTATEGENERATOR": get_val("__VIEWSTATEGENERATOR"),
        "__EVENTVALIDATION": get_val("__EVENTVALIDATION"),  # page mein na ho to khali rahega, theek hai
        "HeadLoginView$logId$UserName": username,
        "HeadLoginView$logId$Password": password,
    }

    r2 = s.post(LOGIN_ACTION, data=payload, allow_redirects=True)

    # Login success check: STEDA login fail pe wapas Login.aspx pe hi rakhta hai
    if "Login.aspx" in r2.url and "ReturnUrl" in r2.url:
        return jsonify({"success": False, "message": "Login failed - check username/password"}), 401

    # User ka session token banao aur cookie jar save karo
    token = str(id(s)) + username
    USER_SESSIONS[token] = s
    session["token"] = token
    return jsonify({"success": True})


@app.route("/api/districts")
def districts():
    token = session.get("token")
    if not token or token not in USER_SESSIONS:
        return jsonify({"success": False, "message": "Not logged in"}), 401
    s = USER_SESSIONS[token]
    r = s.get(f"{BASE}/handlers/Handler_STREAM_District_Performance.ashx")
    return jsonify(r.json())


@app.route("/api/projects")
def projects():
    token = session.get("token")
    if not token or token not in USER_SESSIONS:
        return jsonify({"success": False, "message": "Not logged in"}), 401

    s = USER_SESSIONS[token]
    params = {
        "search": request.args.get("search", ""),
        "entity": request.args.get("entity", ""),
        "district": request.args.get("district", ""),
        "level": request.args.get("level", ""),
        "themeNo": request.args.get("themeNo", ""),
        "status": request.args.get("status", ""),
        "page": request.args.get("page", 1),
        "pageSize": request.args.get("pageSize", 25),
    }
    r = s.get(DASHBOARD_API, params=params)
    return jsonify(r.json())


if __name__ == "__main__":
    import os
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
