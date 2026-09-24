<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>School Academic Management System</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --green:#e3f6e5; --green-dark:#4caf6d;
  --blue:#e5f3fb; --blue-dark:#3d8fc4;
  --saffron:#ff9933; --yellow:#ffd93d; --white:#ffffff;
  --ink:#26313a; --border:#d7e2e0;
}
*{box-sizing:border-box}
body{margin:0;font-family:'Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,var(--green) 0%,var(--blue) 100%);color:var(--ink);min-height:100vh;display:flex;flex-direction:column}
header{background:var(--white);padding:14px 22px;display:flex;justify-content:space-between;align-items:center;border-bottom:4px solid var(--saffron);flex-wrap:wrap;gap:8px}
header h1{font-size:1.15rem;margin:0;color:var(--blue-dark)}
header .who{font-size:.85rem;color:#555}
button{cursor:pointer;border:none;border-radius:7px;padding:8px 14px;font-size:.85rem;font-weight:600;background:var(--blue-dark);color:#fff;transition:.15s}
button:hover{filter:brightness(1.08)}
button.secondary{background:var(--green-dark)}
button.warn{background:#e05555}
button.saffron{background:var(--saffron)}
button.ghost{background:var(--white);color:var(--blue-dark);border:1px solid var(--border)}
main{flex:1;padding:18px;max-width:1150px;margin:0 auto;width:100%}
.card{background:var(--white);border-radius:12px;padding:18px;margin-bottom:16px;box-shadow:0 2px 10px rgba(0,0,0,.06)}
.tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}
.tabs button{background:var(--white);color:var(--ink);border:1px solid var(--border)}
.tabs button.active{background:var(--green-dark);color:#fff;border-color:var(--green-dark)}
input,select,textarea{padding:8px;border:1px solid var(--border);border-radius:6px;font-size:.85rem;width:100%}
input[type=checkbox]{width:auto;height:16px;accent-color:var(--green-dark)}
.chk-label{display:flex;align-items:center;gap:6px;font-weight:400}
label{font-size:.78rem;font-weight:600;color:#555;display:block;margin:8px 0 3px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px}
table{width:100%;border-collapse:collapse;font-size:.85rem;margin-top:10px}
th,td{padding:7px 8px;border-bottom:1px solid var(--border);text-align:left}
th{background:var(--yellow);color:var(--ink)}
tr:hover{background:var(--green)}
.pill{display:inline-block;background:var(--blue);border-radius:20px;padding:2px 9px;font-size:.72rem;margin:1px}
.row-actions button{padding:4px 9px;font-size:.75rem;margin-right:4px}
.login-wrap{flex:1;display:flex;align-items:center;justify-content:center;padding:20px}
.login-box{background:var(--white);padding:32px;border-radius:14px;width:100%;max-width:380px;box-shadow:0 6px 24px rgba(0,0,0,.12);border-top:6px solid var(--saffron)}
.login-box h2{margin-top:0;color:var(--blue-dark)}
.roletoggle{display:flex;gap:8px;margin-bottom:14px}
.roletoggle button{flex:1;background:var(--white);color:var(--ink);border:1px solid var(--border)}
.roletoggle button.active{background:var(--saffron);color:#fff;border-color:var(--saffron)}
.msg{font-size:.8rem;padding:8px;border-radius:6px;margin-top:8px}
.msg.err{background:#fde3e3;color:#b12a2a}
.msg.ok{background:var(--green);color:#256b3d}
footer{background:var(--ink);color:#eee;text-align:center;padding:10px;font-size:.78rem}
.stat{background:var(--green);border-radius:10px;padding:14px;text-align:center}
.stat b{font-size:1.5rem;display:block;color:var(--blue-dark)}
.small{font-size:.75rem;color:#666}
.badge{padding:2px 8px;border-radius:5px;font-size:.72rem;font-weight:700}
.badge.yes{background:var(--green);color:#256b3d}
.badge.no{background:#eee;color:#777}
@media print{header,footer,.tabs,.noprint{display:none!important}main{max-width:100%}}
</style>
</head>
<body>

<div id="loginScreen" class="login-wrap">
  <div class="login-box">
    <h2>🏫 Academic Management System</h2>
    <div class="roletoggle">
      <button id="rAdmin" class="active" onclick="setRole('admin')">Admin</button>
      <button id="rTeacher" onclick="setRole('teacher')">Teacher</button>
    </div>
    <label>Username</label><input id="luser" autocomplete="username" onkeydown="if(event.key==='Enter')doLogin()">
    <label>Password</label><input id="lpass" type="password" autocomplete="current-password" onkeydown="if(event.key==='Enter')doLogin()">
    <button style="width:100%;margin-top:14px" onclick="doLogin()">Login</button>
    <div id="loginMsg"></div>
    <div class="small" style="margin-top:10px">Default admin login: <b>admin / admin123</b> (change it after first login)</div>
  </div>
</div>

<div id="app" style="display:none;flex-direction:column;flex:1">
  <header>
    <h1>🏫 School Academic Management System</h1>
    <div class="who"><span id="whoAmI"></span> &nbsp; <button class="ghost" onclick="logout()">Logout</button></div>
  </header>
  <main>
    <div class="tabs" id="tabs"></div>
    <div id="content"></div>
  </main>
</div>

<footer>Concept, Designed and Developed by <b>Ghanu Sharma, Namchi-Sikkim</b> &nbsp;|&nbsp; 9832095992 &nbsp;|&nbsp; ghanusharma805@gmail.com</footer>

<script>
const CLASSES = ['Nursery','LKG','UKG','I','II','III','IV','V','VI','VII','VIII'];
const SECTIONED = ['Nursery','LKG','UKG','I','II','III','IV','V'];
function classSections(c){ return SECTIONED.includes(c) ? ['A','B'] : ['']; }
function classLabel(c,s){ return s ? c+'-'+s : c; }
function assessLabel(a){ return a==='HYE'?'Half Yearly Examination':a==='YE'?'Yearly Examination':a; }

let DB = null;
function load(){
  let raw = localStorage.getItem('sams_data_v1');
  if(raw){ DB = JSON.parse(raw); }
  else { DB = { admin:{username:'admin',password:'admin123'}, subjects:['English','Mathematics','Science','Social Science','Hindi/Nepali'], teachers:[], students:[] }; }
  if(!DB.assessments) DB.assessments = ['UT1','UT2','UT3','UT4','HYE','UT5','UT6','UT7','YE'];
  if(!DB.fullMarks) DB.fullMarks = {UT1:20,UT2:20,UT3:20,UT4:20,HYE:80,UT5:20,UT6:20,UT7:20,YE:80};
  if(!DB.homework) DB.homework = [];
  if(!DB.marks) DB.marks = [];
  if(!DB.schedule) DB.schedule = {};
  DB.assessments.forEach(a=>{
    const cur = DB.schedule[a];
    if(!cur){ DB.schedule[a] = {}; }
    else if(typeof cur.testDate !== 'undefined' || typeof cur.deadline !== 'undefined'){
      const flat = {testDate: cur.testDate||'', deadline: cur.deadline||''};
      const migrated = {};
      DB.subjects.forEach(s=>{ migrated[s] = {testDate: flat.testDate, deadline: flat.deadline}; });
      DB.schedule[a] = migrated;
    }
  });
  DB.teachers.forEach(t=>{
    if(!t.assignments){
      t.assignments = (t.subjects||[]).map(s=>({subject:s, classes:[]}));
      delete t.subjects;
    }
  });
  save();
}
function save(){ localStorage.setItem('sams_data_v1', JSON.stringify(DB)); }
function uid(){ return Date.now().toString(36)+Math.random().toString(36).slice(2,6); }
function myTeacher(){ return DB.teachers.find(t=>t.id===SESSION.id); }

let SESSION = JSON.parse(sessionStorage.getItem('sams_session')||'null');
let role='admin', activeTab='dashboard', selClass=CLASSES[0], selSection='';
let hClass=CLASSES[0], hSection='', expandedHwId=null;
let mAssess='UT1', mClass=CLASSES[0], mSection='', mSubject='';
let expandedSchedAssess=null;

function setRole(r){ role=r; document.getElementById('rAdmin').classList.toggle('active',r==='admin'); document.getElementById('rTeacher').classList.toggle('active',r==='teacher'); }

function doLogin(){
  const u=document.getElementById('luser').value.trim(), p=document.getElementById('lpass').value;
  const msg=document.getElementById('loginMsg'); msg.innerHTML='';
  if(role==='admin'){
    if(u===DB.admin.username && p===DB.admin.password){ SESSION={role:'admin',name:'Administrator'}; afterLogin(); }
    else msg.innerHTML='<div class="msg err">Incorrect admin username or password.</div>';
  } else {
    const t = DB.teachers.find(t=>t.username===u && t.password===p);
    if(t){ SESSION={role:'teacher',id:t.id,name:t.name}; afterLogin(); }
    else msg.innerHTML='<div class="msg err">Incorrect teacher username or password.</div>';
  }
}
function afterLogin(){
  sessionStorage.setItem('sams_session', JSON.stringify(SESSION));
  document.getElementById('loginScreen').style.display='none';
  document.getElementById('app').style.display='flex';
  document.getElementById('whoAmI').textContent = (SESSION.role==='admin'?'Admin':'Teacher')+': '+SESSION.name;
  activeTab = SESSION.role==='admin' ? 'dashboard' : 'mydash';
  renderTabs(); renderTab();
}
function logout(){ sessionStorage.removeItem('sams_session'); location.reload(); }

function renderTabs(){
  const tabs = SESSION.role==='admin'
    ? [['dashboard','Dashboard'],['classes','Classes & Subjects'],['teachers','Teachers'],['students','Students'],['homework','Homework'],['marks','Marks Entry'],['export','Export']]
    : [['mydash','My Info'],['homework','Homework'],['marks','Marks Entry']];
  document.getElementById('tabs').innerHTML = tabs.map(([k,l])=>`<button class="${activeTab===k?'active':''}" onclick="goTab('${k}')">${l}</button>`).join('');
}
function goTab(k){ activeTab=k; renderTabs(); renderTab(); }

function renderTab(){
  const c = document.getElementById('content');
  if(activeTab==='dashboard') c.innerHTML = dashboardHtml();
  else if(activeTab==='classes') c.innerHTML = classesHtml();
  else if(activeTab==='teachers') c.innerHTML = teachersHtml();
  else if(activeTab==='students') c.innerHTML = studentsHtml();
  else if(activeTab==='homework') c.innerHTML = homeworkHtml();
  else if(activeTab==='marks') c.innerHTML = marksHtml();
  else if(activeTab==='export') c.innerHTML = exportHtml();
  else if(activeTab==='mydash') c.innerHTML = teacherDashHtml();
}

/* ---------- DASHBOARD ---------- */
function dashboardHtml(){
  return `<div class="grid">
    <div class="card stat"><b>${CLASSES.length}</b>Classes</div>
    <div class="card stat"><b>${DB.subjects.length}</b>Subjects</div>
    <div class="card stat"><b>${DB.teachers.length}</b>Teachers</div>
    <div class="card stat"><b>${DB.students.length}</b>Students</div>
    <div class="card stat"><b>${DB.homework.length}</b>Homework Entries</div>
    <div class="card stat"><b>${DB.marks.length}</b>Marks Records</div>
  </div>
  <div class="card"><b>Welcome, Admin.</b> Use the tabs above to manage classes/subjects, teacher and student records, homework, marks entry, and export data.</div>`;
}

/* ---------- CLASSES & SUBJECTS ---------- */
function classesHtml(){
  return `<div class="card">
    <h3>Class Structure (fixed)</h3>
    <table><tr><th>Class</th><th>Sections</th></tr>
    ${CLASSES.map(c=>`<tr><td>${c}</td><td>${classSections(c).filter(Boolean).join(', ')||'— (no sections)'}</td></tr>`).join('')}
    </table>
  </div>
  <div class="card">
    <h3>Subjects</h3>
    <div style="display:flex;gap:8px;margin-bottom:10px">
      <input id="newSubj" placeholder="Add new subject name">
      <button onclick="addSubject()">Add</button>
    </div>
    <table><tr><th>#</th><th>Subject</th><th>Action</th></tr>
    ${DB.subjects.map((s,i)=>`<tr><td>${i+1}</td><td>${s}</td><td class="row-actions">
      <button class="secondary" onclick="editSubject(${i})">Edit</button>
      <button class="warn" onclick="removeSubject(${i})">Remove</button></td></tr>`).join('')}
    </table>
  </div>
  <div class="card">
    <h3>Test &amp; Examination Schedule</h3>
    <p class="small">For each assessment, set the Test/Exam Date and the Deadline to submit marks — separately per subject, since subjects are usually tested on different days within the same UT/Exam period.</p>
    <table><tr><th>Assessment</th><th>Full Marks</th><th>Action</th></tr>
    ${DB.assessments.map(a=>`<tr><td>${assessLabel(a)}</td><td>${DB.fullMarks[a]}</td><td class="row-actions">
      <button onclick="toggleSchedAssess('${a}')">${expandedSchedAssess===a?'Hide Schedule':'Manage Schedule'}</button>
      <button class="secondary" onclick="editFullMarks('${a}')">Edit Full Marks</button>
    </td></tr>${expandedSchedAssess===a?`<tr><td colspan="3">${scheduleTable(a)}</td></tr>`:''}`).join('')}
    </table>
    <p class="small">Marks entry locks for teachers, per subject, once that subject's deadline for the selected assessment has passed (Admin can always enter/edit).</p>
  </div>`;
}
function toggleSchedAssess(a){ expandedSchedAssess = expandedSchedAssess===a ? null : a; renderTab(); }
function scheduleTable(a){
  const sch = DB.schedule[a] || {};
  if(DB.subjects.length===0) return '<p class="small">Add subjects first to set their exam schedule.</p>';
  return `<div style="padding:10px;background:var(--blue);border-radius:8px">
    <table><tr><th>Subject</th><th>Test/Exam Date</th><th>Deadline of Marks Submission</th></tr>
    ${DB.subjects.map(s=>{
      const e = sch[s] || {testDate:'',deadline:''};
      return `<tr><td>${s}</td>
      <td><input type="date" value="${e.testDate}" onchange="setSchedule('${a}','${s.replace(/'/g,"\\'")}','testDate',this.value)"></td>
      <td><input type="date" value="${e.deadline}" onchange="setSchedule('${a}','${s.replace(/'/g,"\\'")}','deadline',this.value)"></td></tr>`;
    }).join('')}
    </table>
  </div>`;
}
function addSubject(){
  const v=document.getElementById('newSubj').value.trim();
  if(!v) return alert('Enter a subject name.');
  DB.subjects.push(v); save(); renderTab();
}
function editSubject(i){
  const v=prompt('Edit subject name:', DB.subjects[i]);
  if(v && v.trim()){ DB.subjects[i]=v.trim(); save(); renderTab(); }
}
function removeSubject(i){ if(confirm('Remove this subject?')){ DB.subjects.splice(i,1); save(); renderTab(); } }
function editFullMarks(a){
  const v=prompt(`Full marks for ${assessLabel(a)}:`, DB.fullMarks[a]);
  const n=parseFloat(v);
  if(!v||isNaN(n)||n<=0) return;
  DB.fullMarks[a]=n; save(); renderTab();
}
function setSchedule(a,subject,field,val){
  if(!DB.schedule[a]) DB.schedule[a]={};
  if(!DB.schedule[a][subject]) DB.schedule[a][subject]={testDate:'',deadline:''};
  DB.schedule[a][subject][field]=val; save(); renderTab();
}
function isDeadlinePassed(a,subject){
  const e = DB.schedule[a] && DB.schedule[a][subject];
  if(!e || !e.deadline) return false;
  const today=new Date().toISOString().slice(0,10);
  return today > e.deadline;
}

/* ---------- TEACHERS ---------- */
function teachersHtml(){
  const classOpts = CLASSES.flatMap(c=>classSections(c).map(s=>`<option value="${classLabel(c,s)}">${classLabel(c,s)}</option>`)).join('');
  const subjRows = DB.subjects.map(s=>`<tr>
    <td style="text-align:center"><input type="checkbox" class="tSubjChk" value="${s}"></td>
    <td>${s}</td>
    <td><select multiple class="tSubjClass" size="3" style="min-width:160px">${classOpts}</select></td>
  </tr>`).join('');
  return `<div class="card">
    <h3>Add Teacher (single)</h3>
    <div class="grid">
      <div><label>Name</label><input id="tName"></div>
      <div><label>Class Teacher?</label><select id="tIsCT" onchange="document.getElementById('tCTof').style.display=this.value==='Yes'?'block':'none'"><option>No</option><option>Yes</option></select></div>
      <div id="tCTof" style="display:none"><label>Class Teacher Of</label><select id="tCTclass">${classOpts}</select></div>
      <div><label>Username</label><input id="tUser"></div>
      <div><label>Password</label><input id="tPass"></div>
    </div>
    <label>Subjects Taught</label>
    <p class="small">Tick a subject, then select the class(es) it's taught in for this teacher (hold Ctrl/Cmd to pick more than one).</p>
    <table id="teacherSubjTable"><tr><th>Check box</th><th>Subject</th><th>Class</th></tr>${subjRows}</table>
    <button style="margin-top:12px" onclick="addTeacher()">Add Teacher</button>
  </div>
  <div class="card">
    <h3>Bulk Add Teachers (copy/paste)</h3>
    <p class="small">One teacher per line: <b>Name, ClassTeacher(Y/N), ClassTeacherOf, Subjects(;), Username, Password</b><br>Example: <i>Anita Rai, Y, VI-A, Mathematics;Science, anita, pass123</i></p>
    <textarea id="tBulk" rows="5" placeholder="Paste rows here..."></textarea>
    <button style="margin-top:8px" onclick="bulkAddTeachers()">Import Teachers</button>
  </div>
  <div class="card">
    <h3>Teacher Records (${DB.teachers.length})</h3>
    <table><tr><th>Name</th><th>Class Teacher</th><th>Subjects</th><th>Username</th><th>Action</th></tr>
    ${DB.teachers.map(t=>`<tr>
      <td>${t.name}</td>
      <td>${t.isClassTeacher?`<span class="badge yes">${t.classTeacherOf}</span>`:'<span class="badge no">No</span>'}</td>
      <td>${t.assignments.map(a=>`<span class="pill">${a.subject}${a.classes.length?': '+a.classes.join(', '):''}</span>`).join('') || '<span class="small">None</span>'}</td>
      <td>${t.username}</td>
      <td class="row-actions"><button class="secondary" onclick="editTeacher('${t.id}')">Edit</button><button class="warn" onclick="removeTeacher('${t.id}')">Remove</button></td>
    </tr>`).join('') || '<tr><td colspan="5" class="small">No teachers added yet.</td></tr>'}
    </table>
  </div>`;
}
function addTeacher(){
  const name=document.getElementById('tName').value.trim();
  const user=document.getElementById('tUser').value.trim();
  const pass=document.getElementById('tPass').value.trim();
  const isCT=document.getElementById('tIsCT').value==='Yes';
  const ctOf=isCT?document.getElementById('tCTclass').value:'';
  const assignments=[];
  document.querySelectorAll('#teacherSubjTable tr').forEach(tr=>{
    const cb=tr.querySelector('.tSubjChk');
    if(cb && cb.checked){
      const sel=tr.querySelector('.tSubjClass');
      const classes = sel ? [...sel.selectedOptions].map(o=>o.value) : [];
      assignments.push({subject:cb.value, classes});
    }
  });
  if(!name||!user||!pass) return alert('Name, username and password are required.');
  if(DB.teachers.some(t=>t.username===user)) return alert('Username already exists.');
  DB.teachers.push({id:uid(),name,isClassTeacher:isCT,classTeacherOf:ctOf,assignments,username:user,password:pass});
  save(); renderTab();
}
function bulkAddTeachers(){
  const lines=document.getElementById('tBulk').value.split('\n').map(l=>l.trim()).filter(Boolean);
  let added=0, skipped=0;
  lines.forEach(line=>{
    const parts=line.split(',').map(p=>p.trim());
    if(parts.length<5){ skipped++; return; }
    const [name,ctFlag,ctOf,subjStr,user,pass]=parts.length===6?parts:[parts[0],parts[1],parts[2],parts[3],parts[4],parts[5]||'changeme'];
    if(!name||!user){ skipped++; return; }
    if(DB.teachers.some(t=>t.username===user)){ skipped++; return; }
    const assignments = (subjStr||'').split(';').map(s=>s.trim()).filter(Boolean).map(s=>({subject:s, classes:[]}));
    DB.teachers.push({id:uid(),name,isClassTeacher:/^y/i.test(ctFlag),classTeacherOf:ctOf||'',assignments,username:user,password:pass||'changeme'});
    added++;
  });
  alert(`Imported ${added} teacher(s). Skipped ${skipped} (missing/duplicate). Class assignments per subject were left blank for bulk-imported teachers — set them via Edit.`);
  save(); renderTab();
}
function editTeacher(id){
  const t=DB.teachers.find(x=>x.id===id); if(!t) return;
  const name=prompt('Name:',t.name); if(name===null) return;
  const cur = t.assignments.map(a=>`${a.subject}:${a.classes.join(',')}`).join('; ');
  const raw = prompt('Subjects & Classes — format Subject:Class,Class; Subject2:Class3', cur);
  t.name=name.trim()||t.name;
  if(raw!==null){
    t.assignments = raw.split(';').map(p=>p.trim()).filter(Boolean).map(p=>{
      const [subj,classesStr]=p.split(':');
      return {subject:(subj||'').trim(), classes:(classesStr||'').split(',').map(c=>c.trim()).filter(Boolean)};
    }).filter(a=>a.subject);
  }
  save(); renderTab();
}
function removeTeacher(id){ if(confirm('Remove this teacher?')){ DB.teachers=DB.teachers.filter(t=>t.id!==id); save(); renderTab(); } }

/* ---------- STUDENTS ---------- */
function studentsHtml(){
  const classOpts = CLASSES.map(c=>`<option value="${c}" ${c===selClass?'selected':''}>${c}</option>`).join('');
  const secs = classSections(selClass);
  const secOpts = secs.map(s=>`<option value="${s}" ${s===selSection?'selected':''}>${s||'(none)'}</option>`).join('');
  if(!secs.includes(selSection)) selSection = secs[0];
  const list = DB.students.filter(s=>s.class===selClass && s.section===selSection).sort((a,b)=>(+a.roll)-(+b.roll));
  return `<div class="card">
    <h3>Select Class</h3>
    <div class="grid">
      <div><label>Class</label><select onchange="selClass=this.value; renderTab()">${classOpts}</select></div>
      ${secs[0]?`<div><label>Section</label><select onchange="selSection=this.value; renderTab()">${secOpts}</select></div>`:''}
    </div>
  </div>
  <div class="card">
    <h3>Add Student — ${classLabel(selClass,selSection)}</h3>
    <div class="grid">
      <div><label>Roll No</label><input id="sRoll" type="number"></div>
      <div><label>Name</label><input id="sName"></div>
    </div>
    <button style="margin-top:10px" onclick="addStudent()">Add Student</button>
  </div>
  <div class="card">
    <h3>Bulk Add Students (copy/paste)</h3>
    <p class="small">Paste rows as <b>Roll, Name</b> (one per line) for ${classLabel(selClass,selSection)} — works directly from Excel copy.</p>
    <textarea id="sBulk" rows="5" placeholder="1, Rahul Sharma&#10;2, Priya Rai"></textarea>
    <button style="margin-top:8px" onclick="bulkAddStudents()">Import Students</button>
  </div>
  <div class="card">
    <h3>Student List — ${classLabel(selClass,selSection)} (${list.length})</h3>
    <table><tr><th>Roll</th><th>Name</th><th>Action</th></tr>
    ${list.map(s=>`<tr><td>${s.roll}</td><td>${s.name}</td><td class="row-actions">
      <button class="secondary" onclick="editStudent('${s.id}')">Edit</button>
      <button class="warn" onclick="removeStudent('${s.id}')">Remove</button></td></tr>`).join('') || '<tr><td colspan="3" class="small">No students in this class/section yet.</td></tr>'}
    </table>
  </div>`;
}
function addStudent(){
  const roll=document.getElementById('sRoll').value.trim();
  const name=document.getElementById('sName').value.trim();
  if(!roll||!name) return alert('Roll no. and name are required.');
  if(DB.students.some(s=>s.class===selClass&&s.section===selSection&&s.roll===roll)) return alert('Roll number already exists in this class/section.');
  DB.students.push({id:uid(),class:selClass,section:selSection,roll,name});
  save(); renderTab();
}
function bulkAddStudents(){
  const lines=document.getElementById('sBulk').value.split('\n').map(l=>l.trim()).filter(Boolean);
  let added=0, skipped=0;
  lines.forEach(line=>{
    const parts=line.split(/,|\t/).map(p=>p.trim());
    const roll=parts[0], name=parts[1];
    if(!roll||!name){ skipped++; return; }
    if(DB.students.some(s=>s.class===selClass&&s.section===selSection&&s.roll===roll)){ skipped++; return; }
    DB.students.push({id:uid(),class:selClass,section:selSection,roll,name});
    added++;
  });
  alert(`Imported ${added} student(s). Skipped ${skipped} (missing/duplicate roll).`);
  save(); renderTab();
}
function editStudent(id){
  const s=DB.students.find(x=>x.id===id); if(!s) return;
  const roll=prompt('Roll No:',s.roll); if(roll===null) return;
  const name=prompt('Name:',s.name); if(name===null) return;
  s.roll=roll.trim()||s.roll; s.name=name.trim()||s.name;
  save(); renderTab();
}
function removeStudent(id){ if(confirm('Remove this student?')){ DB.students=DB.students.filter(s=>s.id!==id); save(); renderTab(); } }

/* ---------- HOMEWORK ---------- */
function homeworkHtml(){
  const classOpts = CLASSES.map(c=>`<option value="${c}" ${c===hClass?'selected':''}>${c}</option>`).join('');
  const secs = classSections(hClass);
  if(!secs.includes(hSection)) hSection = secs[0];
  const secOpts = secs.map(s=>`<option value="${s}" ${s===hSection?'selected':''}>${s||'(none)'}</option>`).join('');
  const availSubjects = SESSION.role==='teacher' ? (myTeacher()?myTeacher().assignments.map(a=>a.subject):[]) : DB.subjects;
  const subjOpts = availSubjects.map(s=>`<option>${s}</option>`).join('');
  const list = DB.homework.filter(h => SESSION.role==='admin' || h.teacherId===SESSION.id).sort((a,b)=> b.givenDate.localeCompare(a.givenDate));
  return `<div class="card">
    <h3>Add Homework</h3>
    ${availSubjects.length===0?'<p class="small">No subjects assigned to you yet — ask the Admin to assign subjects.</p>':`
    <div class="grid">
      <div><label>Class</label><select id="hClassSel" onchange="hClass=this.value; renderTab()">${classOpts}</select></div>
      ${secs[0]?`<div><label>Section</label><select id="hSecSel" onchange="hSection=this.value; renderTab()">${secOpts}</select></div>`:''}
      <div><label>Subject</label><select id="hSubjSel">${subjOpts}</select></div>
      <div><label>Given Date</label><input id="hGiven" type="date" value="${new Date().toISOString().slice(0,10)}"></div>
      <div><label>Due Date</label><input id="hDue" type="date"></div>
    </div>
    <label>Homework Details</label>
    <textarea id="hDetails" rows="2" placeholder="e.g. Exercise 3.2, Q1-10"></textarea>
    <button style="margin-top:10px" onclick="addHomework()">Add Homework</button>`}
  </div>
  <div class="card">
    <h3>Homework Records (${list.length})</h3>
    <table><tr><th>Given</th><th>Due</th><th>Class</th><th>Subject</th><th>Details</th><th>Submitted</th><th>Action</th></tr>
    ${list.map(h=>{
      const students=DB.students.filter(s=>s.class===h.class&&s.section===h.section);
      const subCount=students.filter(s=>h.submissions[s.id]==='Submitted').length;
      const row = `<tr>
        <td>${h.givenDate}</td><td>${h.dueDate||'-'}</td><td>${classLabel(h.class,h.section)}</td><td>${h.subject}</td><td>${h.details}</td>
        <td>${subCount}/${students.length}</td>
        <td class="row-actions">
          <button onclick="toggleExpandHw('${h.id}')">Track</button>
          <button class="warn" onclick="removeHomework('${h.id}')">Remove</button>
        </td>
      </tr>`;
      const expand = expandedHwId===h.id ? `<tr><td colspan="7">${hwSubmissionPanel(h)}</td></tr>` : '';
      return row+expand;
    }).join('') || '<tr><td colspan="7" class="small">No homework recorded yet.</td></tr>'}
    </table>
  </div>`;
}
function hwSubmissionPanel(h){
  const students=DB.students.filter(s=>s.class===h.class&&s.section===h.section).sort((a,b)=>(+a.roll)-(+b.roll));
  return `<div style="padding:10px;background:var(--green);border-radius:8px">
    <b>Mark Submission — ${classLabel(h.class,h.section)} ${h.subject}</b>
    <table><tr><th>Roll</th><th>Name</th><th>Status</th><th>Action</th></tr>
    ${students.map(s=>{
      const st=h.submissions[s.id]||'Not Submitted';
      return `<tr><td>${s.roll}</td><td>${s.name}</td><td>${st==='Submitted'?'<span class="badge yes">Submitted</span>':'<span class="badge no">Not Submitted</span>'}</td>
      <td><button class="secondary" onclick="setSubmission('${h.id}','${s.id}','Submitted')">Submitted</button> <button class="ghost" onclick="setSubmission('${h.id}','${s.id}','Not Submitted')">Not Submitted</button></td></tr>`;
    }).join('') || '<tr><td colspan="4" class="small">No students in this class/section.</td></tr>'}
    </table>
    <button class="ghost" style="margin-top:8px" onclick="expandedHwId=null; renderTab()">Close</button>
  </div>`;
}
function addHomework(){
  const cls=document.getElementById('hClassSel').value;
  const secEl=document.getElementById('hSecSel'); const sec=secEl?secEl.value:'';
  const subj=document.getElementById('hSubjSel').value;
  const given=document.getElementById('hGiven').value;
  const due=document.getElementById('hDue').value;
  const details=document.getElementById('hDetails').value.trim();
  if(!given||!details) return alert('Given date and homework details are required.');
  DB.homework.push({id:uid(),class:cls,section:sec,subject:subj,teacherId:SESSION.role==='teacher'?SESSION.id:'admin',teacherName:SESSION.name,givenDate:given,dueDate:due,details,submissions:{}});
  save(); renderTab();
}
function setSubmission(hwId,studentId,status){
  const h=DB.homework.find(x=>x.id===hwId); if(!h) return;
  h.submissions[studentId]=status; save(); renderTab();
}
function removeHomework(id){ if(confirm('Remove this homework entry?')){ DB.homework=DB.homework.filter(h=>h.id!==id); if(expandedHwId===id) expandedHwId=null; save(); renderTab(); } }
function toggleExpandHw(id){ expandedHwId = expandedHwId===id?null:id; renderTab(); }

/* ---------- MARKS ENTRY ---------- */
function marksHtml(){
  const availSubjects = SESSION.role==='teacher' ? (myTeacher()?myTeacher().assignments.map(a=>a.subject):[]) : DB.subjects;
  if(availSubjects.length===0) return '<div class="card"><p class="small">No subjects assigned to you yet — ask the Admin to assign subjects.</p></div>';
  const classOpts = CLASSES.map(c=>`<option value="${c}" ${c===mClass?'selected':''}>${c}</option>`).join('');
  const secs = classSections(mClass);
  if(!secs.includes(mSection)) mSection = secs[0];
  const secOpts = secs.map(s=>`<option value="${s}" ${s===mSection?'selected':''}>${s||'(none)'}</option>`).join('');
  if(!availSubjects.includes(mSubject)) mSubject = availSubjects[0];
  const subjOpts = availSubjects.map(s=>`<option ${s===mSubject?'selected':''}>${s}</option>`).join('');
  const assessOpts = DB.assessments.map(a=>`<option value="${a}" ${a===mAssess?'selected':''}>${assessLabel(a)}</option>`).join('');
  const fm = DB.fullMarks[mAssess];
  const sch = (DB.schedule[mAssess] && DB.schedule[mAssess][mSubject]) || {testDate:'',deadline:''};
  const locked = isDeadlinePassed(mAssess,mSubject) && SESSION.role==='teacher';
  const students = DB.students.filter(s=>s.class===mClass&&s.section===mSection).sort((a,b)=>(+a.roll)-(+b.roll));
  const existing = {}; DB.marks.filter(m=>m.class===mClass&&m.section===mSection&&m.subject===mSubject&&m.assessment===mAssess).forEach(m=>existing[m.studentId]=m);
  return `<div class="card">
    <h3>Marks Entry</h3>
    <div class="grid">
      <div><label>Assessment</label><select onchange="mAssess=this.value; renderTab()">${assessOpts}</select></div>
      <div><label>Class</label><select onchange="mClass=this.value; renderTab()">${classOpts}</select></div>
      ${secs[0]?`<div><label>Section</label><select onchange="mSection=this.value; renderTab()">${secOpts}</select></div>`:''}
      <div><label>Subject</label><select onchange="mSubject=this.value; renderTab()">${subjOpts}</select></div>
    </div>
    <p style="margin-top:10px"><b>Full Marks for ${assessLabel(mAssess)}: ${fm}</b> &nbsp; <span class="small">(entries above ${fm} or below 0 will show an error)</span></p>
    <p class="small">📅 Test/Exam Date (${mSubject}): <b>${sch.testDate||'Not set'}</b> &nbsp;|&nbsp; ⏰ Marks Submission Deadline: <b>${sch.deadline||'Not set'}</b></p>
    ${isDeadlinePassed(mAssess,mSubject) ? `<div class="msg err">⚠️ The marks submission deadline for ${mSubject} — ${assessLabel(mAssess)} has passed${SESSION.role==='teacher'?' — entry is locked. Contact Admin.':' (Admin can still enter/edit marks).'}</div>` : ''}
  </div>
  <div class="card">
    <h3>${classLabel(mClass,mSection)} — ${mSubject} — ${assessLabel(mAssess)}</h3>
    ${students.length===0?'<p class="small">No students in this class/section.</p>':`
    <table><tr><th>Roll</th><th>Name</th><th>Marks (/${fm})</th><th>Absent</th><th>Status</th></tr>
    ${students.map(s=>{
      const ex = existing[s.id];
      const absent = ex && ex.status==='Absent';
      const val = ex && !absent ? ex.marks : '';
      return `<tr>
        <td>${s.roll}</td><td>${s.name}</td>
        <td><input type="number" id="mk_${s.id}" value="${val}" ${absent||locked?'disabled':''} style="width:80px" oninput="validateMarkInput('${s.id}',${fm})"></td>
        <td style="text-align:center"><input type="checkbox" id="ab_${s.id}" ${absent?'checked':''} ${locked?'disabled':''} onchange="toggleAbsent('${s.id}',${fm})"></td>
        <td id="st_${s.id}">${markStatusBadge(ex,fm)}</td>
      </tr>`;
    }).join('')}
    </table>
    <button style="margin-top:12px" onclick="saveMarks('${mClass}','${mSection}','${mSubject}','${mAssess}',${fm})" ${locked?'disabled':''}>Save Marks</button>
    `}
  </div>`;
}
function markStatusBadge(ex,fm){
  if(!ex) return '<span class="badge no">Not entered</span>';
  if(ex.status==='Absent') return '<span class="badge" style="background:var(--blue);color:#1a5c85">Absent</span>';
  if(ex.marks>fm||ex.marks<0) return '<span class="badge" style="background:#fde3e3;color:#b12a2a">❌ Invalid</span>';
  return '<span class="badge yes">✅ Entered</span>';
}
function validateMarkInput(id,fm){
  const inp=document.getElementById('mk_'+id); const st=document.getElementById('st_'+id);
  const raw=inp.value.trim();
  if(raw===''){ st.innerHTML='<span class="badge no">Not entered</span>'; return; }
  const num=parseFloat(raw);
  if(isNaN(num)||num<0){ st.innerHTML='<span class="badge" style="background:#fde3e3;color:#b12a2a">❌ Invalid marks</span>'; return; }
  if(num>fm){ st.innerHTML=`<span class="badge" style="background:#fde3e3;color:#b12a2a">❌ Cannot exceed ${fm}</span>`; return; }
  st.innerHTML='<span class="badge yes">✅ Entered</span>';
}
function toggleAbsent(id,fm){
  const cb=document.getElementById('ab_'+id); const inp=document.getElementById('mk_'+id); const st=document.getElementById('st_'+id);
  if(cb.checked){ inp.value=''; inp.disabled=true; st.innerHTML='<span class="badge" style="background:var(--blue);color:#1a5c85">Absent</span>'; }
  else{ inp.disabled=false; validateMarkInput(id,fm); }
}
function saveMarks(cls,sec,subj,assess,fm){
  if(isDeadlinePassed(assess,subj) && SESSION.role==='teacher'){ alert('The marks submission deadline for '+subj+' — '+assessLabel(assess)+' has passed. Contact Admin.'); return; }
  const students = DB.students.filter(s=>s.class===cls&&s.section===sec);
  let errors=[], toSave=[];
  students.forEach(s=>{
    const abEl=document.getElementById('ab_'+s.id);
    if(abEl && abEl.checked){ toSave.push({studentId:s.id, marks:null, status:'Absent'}); return; }
    const inp=document.getElementById('mk_'+s.id);
    const raw=inp.value.trim();
    if(raw==='') return;
    const num=parseFloat(raw);
    if(isNaN(num) || num<0){ errors.push(s.name+': Invalid marks.'); return; }
    if(num>fm){ errors.push(s.name+`: Marks cannot exceed ${fm}.`); return; }
    toSave.push({studentId:s.id, marks:num, status:'Entered'});
  });
  if(errors.length){ alert('❌ Cannot save — fix these errors first:\n'+errors.join('\n')); return; }
  toSave.forEach(rec=>{
    const idx = DB.marks.findIndex(m=>m.class===cls&&m.section===sec&&m.subject===subj&&m.assessment===assess&&m.studentId===rec.studentId);
    const entry = {id: idx>=0?DB.marks[idx].id:uid(), class:cls, section:sec, subject:subj, assessment:assess, studentId:rec.studentId, marks:rec.marks, status:rec.status, enteredBy:SESSION.name, date:new Date().toISOString().slice(0,10)};
    if(idx>=0) DB.marks[idx]=entry; else DB.marks.push(entry);
  });
  save();
  alert('✅ Marks saved successfully.');
  renderTab();
}

/* ---------- EXPORT ---------- */
function exportHtml(){
  return `<div class="card">
    <h3>Export Data</h3>
    <p class="small">Download full records as Excel, or use Print for a PDF (your browser's Print dialog → "Save as PDF").</p>
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px">
      <button class="secondary" onclick="exportExcel('teachers')">Teachers → Excel</button>
      <button class="secondary" onclick="exportExcel('students')">Students → Excel</button>
      <button class="secondary" onclick="exportExcel('homework')">Homework → Excel</button>
      <button class="secondary" onclick="exportExcel('marks')">Marks → Excel</button>
      <button class="saffron" onclick="window.print()">Print / Save as PDF</button>
    </div>
  </div>`;
}
function exportExcel(kind){
  let rows;
  if(kind==='teachers') rows = DB.teachers.map(t=>({Name:t.name,'Class Teacher':t.isClassTeacher?'Yes':'No','Class Teacher Of':t.classTeacherOf,'Subjects & Classes':t.assignments.map(a=>a.subject+(a.classes.length?' ('+a.classes.join(', ')+')':'')).join('; '),Username:t.username}));
  else if(kind==='students') rows = DB.students.map(s=>({Class:s.class,Section:s.section,Roll:s.roll,Name:s.name}));
  else if(kind==='homework') rows = DB.homework.map(h=>({Class:classLabel(h.class,h.section),Subject:h.subject,'Given Date':h.givenDate,'Due Date':h.dueDate,Details:h.details,Teacher:h.teacherName}));
  else if(kind==='marks') rows = DB.marks.map(m=>{ const st=DB.students.find(s=>s.id===m.studentId); return {Class:classLabel(m.class,m.section),Subject:m.subject,Assessment:assessLabel(m.assessment),Roll:st?st.roll:'',Name:st?st.name:'',Marks:m.status==='Absent'?'Absent':m.marks,'Entered By':m.enteredBy}; });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, kind);
  XLSX.writeFile(wb, kind+'_export.xlsx');
}

/* ---------- TEACHER SELF VIEW ---------- */
function teacherDashHtml(){
  const t = myTeacher();
  if(!t) return '<div class="card">Record not found.</div>';
  let classList='';
  if(t.isClassTeacher){
    const [c,s] = (t.classTeacherOf||'').split('-');
    const list = DB.students.filter(st=>st.class===c && (st.section===s || (!s && !st.section))).sort((a,b)=>(+a.roll)-(+b.roll));
    classList = `<div class="card"><h3>My Class: ${t.classTeacherOf}</h3>
      <table><tr><th>Roll</th><th>Name</th></tr>
      ${list.map(st=>`<tr><td>${st.roll}</td><td>${st.name}</td></tr>`).join('')||'<tr><td colspan="2" class="small">No students recorded yet.</td></tr>'}
      </table></div>`;
  }
  return `<div class="card">
    <h3>Welcome, ${t.name}</h3>
    <p><b>Class Teacher:</b> ${t.isClassTeacher?`Yes — ${t.classTeacherOf}`:'No'}</p>
    <p><b>Subjects &amp; Classes Taught:</b> ${t.assignments.map(a=>`<span class="pill">${a.subject}${a.classes.length?': '+a.classes.join(', '):''}</span>`).join('')||'None assigned'}</p>
  </div>${classList}`;
}

/* ---------- INIT ---------- */
load();
if(SESSION){ afterLogin(); }
</script>
</body>
</html>
