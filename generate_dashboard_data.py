"""
SCTS Student Failure Analysis - Builds dashboard.html with data embedded inline.
Run this script to regenerate the dashboard whenever Excel files change.
"""

import json, openpyxl, os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FILES = {
    'PRE_MID':  {'path': os.path.join(BASE_DIR, 'CBSE KA TN 6-9 PRE MID TERM Marks Entry Format (1) (8).xlsx'), 'label': 'Pre-Midterm',  'pass_mark': 14, 'total_theory': 40},
    'MID_TERM': {'path': os.path.join(BASE_DIR, 'CBSE KA TN MID TERM Marks Entry Format (5).xlsx'),             'label': 'Mid-Term',     'pass_mark': 28, 'total_theory': 80},
    'POST_MID': {'path': os.path.join(BASE_DIR, 'CBSE KA TN POST MID Marks Entry Format (3).xlsx'),             'label': 'Post-Midterm', 'pass_mark': 14, 'total_theory': 40},
    'ANNUAL':   {'path': os.path.join(BASE_DIR, 'CBSE KA TN ANNUAL EXAM Marks Entry Format (1).xlsx'),          'label': 'Annual Exam',  'pass_mark': 28, 'total_theory': 80},
}
SUBJECTS_6 = ['FL', 'SL', 'TL', 'Mat', 'GS', 'SS']
SUBJECTS_5 = ['FL', 'SL', 'Mat', 'GS', 'SS']
ORIENTATION_MAP = {'4': 'Techno', '15': 'CBatch', '20': 'IPL', '32': 'MPL'}

def parse_sheet_name(name):
    p = name.split('-')
    if len(p) >= 3:
        return p[0], int(p[1]), p[2], ORIENTATION_MAP.get(p[0], p[0])
    return None, None, None, None

def safe_int(val):
    if val is None or str(val).strip() in ('AB', 'ab', ''): return None
    try: return int(float(str(val)))
    except: return None

def parse_row(row, pass_mark, is_class9):
    try: int(float(str(row[0])))
    except: return None
    if not row[0]: return None
    name = str(row[4]).strip() if row[4] else ''
    if not name: return None
    subjs = SUBJECTS_5 if is_class9 else SUBJECTS_6
    ts = 14 if is_class9 else 15
    tm = {s: safe_int(row[ts+i]) if len(row) > ts+i else None for i, s in enumerate(subjs)}
    if all(m is None for m in tm.values()): return None
    fs = {s: m for s, m in tm.items() if m is None or m < pass_mark}
    return {
        'admin_no': str(row[3]).strip() if row[3] else '',
        'name': name,
        'grade': safe_int(row[6]),
        'section': str(row[7]).strip() if row[7] else '',
        'orientation': str(row[5]).strip() if row[5] else '',
        'theory_marks': tm, 'failed_subjects': fs, 'failed_count': len(fs),
    }

def parse_file(exam_key, info):
    records, wb = [], openpyxl.load_workbook(info['path'], data_only=True)
    for sname in wb.sheetnames:
        prefix, grade, section, orientation = parse_sheet_name(sname)
        if grade is None: continue
        for row in wb[sname].iter_rows(min_row=4, values_only=True):
            if all(c is None for c in row): continue
            rec = parse_row(row, info['pass_mark'], grade == 9)
            if not rec: continue
            rec.update({'orientation': orientation, 'section': section, 'exam_key': exam_key,
                        'exam_label': info['label'], 'pass_mark': info['pass_mark'], 'total_theory': info['total_theory']})
            if rec['grade'] is None: rec['grade'] = grade
            records.append(rec)
    return records

def build_data():
    all_records = []
    for exam_key, info in FILES.items():
        recs = parse_file(exam_key, info)
        print(f'{info["label"]}: {len(recs)} records')
        all_records.extend(recs)
    # Build cross-exam lookup
    cx = {}
    for r in all_records:
        cx.setdefault(r['admin_no'], {})[r['exam_key']] = {
            'theory_marks': r['theory_marks'], 'failed_subjects': r['failed_subjects'],
            'failed_count': r['failed_count'], 'pass_mark': r['pass_mark'],
            'total_theory': r['total_theory'], 'exam_label': r['exam_label'],
        }
    for r in all_records:
        r['exam_history'] = {e: cx[r['admin_no']][e] for e in ['PRE_MID','MID_TERM','POST_MID'] if e in cx.get(r['admin_no'], {})}
    af = [r for r in all_records if r['exam_key']=='ANNUAL' and r['failed_count']>0]
    print(f'Total: {len(all_records)} | Annual failed: {len(af)} | With history: {sum(1 for r in af if r["exam_history"])}')
    return {
        'meta': {
            'grades': sorted({r['grade'] for r in all_records if r['grade']}),
            'exams': [{'key': k, 'label': v['label'], 'pass_mark': v['pass_mark'], 'total_theory': v['total_theory']} for k,v in FILES.items()],
            'sections': sorted({r['section'] for r in all_records if r['section']}),
            'orientations': sorted({r['orientation'] for r in all_records if r['orientation']}),
        },
        'records': all_records
    }

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>SCTS Student Failure Analysis Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--bg:#0f1117;--surface:#1a1d27;--surface2:#222534;--border:#2e3148;--accent:#6c63ff;--text:#e8eaf6;--text-muted:#8890b0;--red:#ff4d6d;--green:#43e97b;--yellow:#f9a825;--blue:#4facfe;--radius:14px}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
.header{background:linear-gradient(135deg,#1a1d27,#0f1117);border-bottom:1px solid var(--border);padding:16px 28px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:100}
.header-icon{width:42px;height:42px;background:linear-gradient(135deg,var(--accent),#a78bfa);border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:20px}
.header h1{font-size:18px;font-weight:700}.header h1 span{color:var(--accent)}
.header-sub{font-size:11px;color:var(--text-muted);margin-top:2px}
.header-right{margin-left:auto;display:flex;gap:8px;align-items:center}
.container{padding:24px 28px;max-width:1560px;margin:0 auto}
.filter-bar{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px 20px;display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end;margin-bottom:24px}
.filter-group{display:flex;flex-direction:column;gap:5px}
.filter-group label{font-size:11px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.8px}
.filter-group select,.filter-group input{background:var(--surface2);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:7px 12px;font-size:13px;font-family:'Inter',sans-serif;cursor:pointer;min-width:125px;outline:none;transition:border-color .2s}
.filter-group select:focus,.filter-group input:focus{border-color:var(--accent)}
.filter-actions{display:flex;gap:8px;margin-left:auto;align-items:flex-end}
.btn{display:inline-flex;align-items:center;gap:6px;padding:8px 16px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;border:none;font-family:'Inter',sans-serif;transition:all .2s;white-space:nowrap}
.btn-primary{background:linear-gradient(135deg,var(--accent),#a78bfa);color:#fff}
.btn-primary:hover{transform:translateY(-1px);box-shadow:0 6px 18px rgba(108,99,255,.4)}
.btn-success{background:linear-gradient(135deg,#43e97b,#38f9d7);color:#000}
.btn-success:hover{transform:translateY(-1px);box-shadow:0 5px 16px rgba(67,233,123,.35)}
.btn-ghost{background:var(--surface2);border:1px solid var(--border);color:var(--text)}
.btn-ghost:hover{border-color:var(--accent);color:var(--accent)}
.btn-xs{padding:4px 10px;font-size:11px;border-radius:6px}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:24px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 22px;position:relative;overflow:hidden;transition:transform .2s,border-color .2s}
.card:hover{transform:translateY(-2px);border-color:var(--accent)}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--radius) var(--radius) 0 0}
.card.blue::before{background:linear-gradient(90deg,var(--blue),#00f2fe)}
.card.red::before{background:linear-gradient(90deg,var(--red),#f9a825)}
.card.green::before{background:linear-gradient(90deg,var(--green),#38f9d7)}
.card.purple::before{background:linear-gradient(90deg,var(--accent),#a78bfa)}
.card.yellow::before{background:linear-gradient(90deg,var(--yellow),#ff6b6b)}
.card-label{font-size:11px;color:var(--text-muted);font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px}
.card-value{font-size:34px;font-weight:800;line-height:1}
.card-sub{font-size:11px;color:var(--text-muted);margin-top:5px}
.card.blue .card-value{color:var(--blue)}.card.red .card-value{color:var(--red)}.card.green .card-value{color:var(--green)}.card.purple .card-value{color:var(--accent)}.card.yellow .card-value{color:var(--yellow)}
.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px}
@media(max-width:900px){.charts-grid{grid-template-columns:1fr}}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px}
.chart-title{font-size:13px;font-weight:700;margin-bottom:3px;display:flex;align-items:center;gap:7px}
.chart-title .badge{font-size:10px;background:var(--surface2);border:1px solid var(--border);border-radius:5px;padding:2px 7px;color:var(--text-muted);font-weight:500}
.chart-sub{font-size:11px;color:var(--text-muted);margin-bottom:16px}
.chart-wrap{position:relative;height:230px}.chart-wrap.tall{height:270px}
.tabs{display:flex;gap:4px;margin-bottom:22px;background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:4px}
.tab{flex:1;text-align:center;padding:8px 14px;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;transition:all .2s;color:var(--text-muted);border:none;background:transparent;font-family:'Inter',sans-serif}
.tab:hover{color:var(--text)}.tab.active{background:var(--accent);color:#fff}
.table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden}
.table-scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:var(--surface2);padding:11px 13px;text-align:left;font-size:10px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--border);white-space:nowrap}
tbody tr.data-row{border-bottom:1px solid #1e2133;transition:background .15s}
tbody tr.data-row:hover{background:rgba(108,99,255,.06)}
tbody tr.comp-row{border-bottom:1px solid var(--border)}
td{padding:10px 13px;vertical-align:middle}
.td-name{font-weight:600}.td-admin{color:var(--text-muted);font-size:11px}
.pill{display:inline-flex;align-items:center;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:700;white-space:nowrap}
.pill-red{background:rgba(255,77,109,.15);color:var(--red)}
.pill-yellow{background:rgba(249,168,37,.15);color:var(--yellow)}
.pill-green{background:rgba(67,233,123,.15);color:var(--green)}
.pill-blue{background:rgba(79,172,254,.15);color:var(--blue)}
.pill-purple{background:rgba(108,99,255,.15);color:var(--accent)}
.subj-box{display:inline-flex;align-items:center;gap:3px;background:rgba(255,77,109,.12);border:1px solid rgba(255,77,109,.25);border-radius:6px;padding:3px 7px;font-size:11px;margin:2px}
.subj-box .sn{color:var(--text-muted);font-weight:600}.subj-box .sm{color:var(--red);font-weight:700}
.subj-box.ab{background:rgba(249,168,37,.1);border-color:rgba(249,168,37,.25)}.subj-box.ab .sm{color:var(--yellow)}
/* COMPARISON TABLE */
.comp-panel{background:#13162080;border-top:1px solid var(--border)}
.comp-inner{padding:14px 16px}
.comp-title{font-size:12px;font-weight:700;color:var(--text-muted);margin-bottom:10px;display:flex;align-items:center;gap:8px}
.comp-table{width:100%;border-collapse:collapse;font-size:12px}
.comp-table th{padding:6px 10px;background:var(--surface2);color:var(--text-muted);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;border:1px solid var(--border);text-align:center}
.comp-table td{padding:7px 10px;border:1px solid #1e2133;text-align:center;font-weight:600}
.comp-table .subj-col{text-align:left;color:var(--text-muted);font-weight:600;background:var(--surface2)}
.c-pass{color:var(--green)}.c-fail{color:var(--red)}.c-ab{color:var(--yellow)}.c-na{color:#444}
.comp-table tr:hover td{background:rgba(108,99,255,.05)}
.toggle-btn{background:rgba(108,99,255,.15);border:1px solid rgba(108,99,255,.3);color:var(--accent);border-radius:6px;padding:3px 9px;font-size:11px;font-weight:600;cursor:pointer;font-family:'Inter',sans-serif;transition:all .2s;white-space:nowrap}
.toggle-btn:hover{background:rgba(108,99,255,.3)}
.toggle-btn.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.pagination{display:flex;align-items:center;gap:7px;padding:12px 18px;border-top:1px solid var(--border)}
.page-info{font-size:12px;color:var(--text-muted);margin-right:auto}
.page-btn{width:30px;height:30px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text);font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .15s;font-family:'Inter',sans-serif}
.page-btn:hover:not(:disabled){border-color:var(--accent);color:var(--accent)}
.page-btn:disabled{opacity:.3;cursor:not-allowed}.page-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.search-wrap{position:relative}
.search-input{background:var(--surface2);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:7px 12px 7px 32px;font-size:13px;font-family:'Inter',sans-serif;outline:none;width:200px;transition:border-color .2s}
.search-input:focus{border-color:var(--accent)}
.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:13px;pointer-events:none}
.section-title{font-size:15px;font-weight:700;margin-bottom:3px}
.section-sub{font-size:12px;color:var(--text-muted);margin-bottom:16px}
.table-header-row{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:12px}
select option{background:#1a1d27}
</style>
</head>
<body>
<div class="header">
  <div class="header-icon">🎓</div>
  <div>
    <h1>SRI CHAITANYA SCHOOL · <span>Failure Analysis</span></h1>
    <div class="header-sub">CBSE KA TN · KR PURAM Branch</div>
  </div>
  <div class="header-right">
    <button class="btn btn-success" onclick="downloadExcel()">⬇️ Download Excel</button>
  </div>
</div>
<div class="container">
  <div class="filter-bar">
    <div class="filter-group"><label>Exam</label>
      <select id="fExam" onchange="applyFilters()">
        <option value="ALL">All Exams</option>
        <option value="PRE_MID">Pre-Midterm</option>
        <option value="MID_TERM">Mid-Term</option>
        <option value="POST_MID">Post-Midterm</option>
        <option value="ANNUAL" selected>Annual Exam</option>
      </select>
    </div>
    <div class="filter-group"><label>Grade</label>
      <select id="fGrade" onchange="applyFilters()">
        <option value="ALL">All Grades</option>
        <option value="6">Grade 6</option><option value="7">Grade 7</option>
        <option value="8">Grade 8</option><option value="9">Grade 9</option>
      </select>
    </div>
    <div class="filter-group"><label>Section</label>
      <select id="fSection" onchange="applyFilters()"><option value="ALL">All Sections</option></select>
    </div>
    <div class="filter-group"><label>Orientation</label>
      <select id="fOrientation" onchange="applyFilters()"><option value="ALL">All Orientations</option></select>
    </div>
    <div class="filter-group"><label>Show Students</label>
      <select id="fStatus" onchange="applyFilters()">
        <option value="FAILED">Failed Only</option>
        <option value="ALL">All Students</option>
      </select>
    </div>
    <div class="filter-actions">
      <div class="search-wrap">
        <span class="search-icon">🔍</span>
        <input type="text" class="search-input" id="fSearch" placeholder="Search student..." oninput="applyFilters()">
      </div>
      <button class="btn btn-ghost" onclick="resetFilters()">Reset</button>
    </div>
  </div>
  <div class="cards-grid">
    <div class="card blue"><div class="card-label">Total Students</div><div class="card-value" id="cTotal">—</div><div class="card-sub" id="cTotalSub">in selected view</div></div>
    <div class="card red"><div class="card-label">Failed Students</div><div class="card-value" id="cFailed">—</div><div class="card-sub" id="cFailedSub">failed ≥1 subject</div></div>
    <div class="card green"><div class="card-label">Pass Rate</div><div class="card-value" id="cPassRate">—</div><div class="card-sub">overall pass %</div></div>
    <div class="card yellow"><div class="card-label">Multi-Subject Fails</div><div class="card-value" id="cMulti">—</div><div class="card-sub">failed 2+ subjects</div></div>
    <div class="card purple"><div class="card-label">Critical (4+ Fails)</div><div class="card-value" id="cCritical">—</div><div class="card-sub">failed 4 or more subjects</div></div>
  </div>
  <div class="tabs">
    <button class="tab active" onclick="switchTab('charts')" id="tab-charts">📊 Charts & Analysis</button>
    <button class="tab" onclick="switchTab('table')" id="tab-table">📋 Failure Details Table</button>
    <button class="tab" onclick="switchTab('trend')" id="tab-trend">📈 Trend Analysis</button>
  </div>
  <!-- CHARTS -->
  <div id="view-charts">
    <div class="charts-grid">
      <div class="chart-card"><div class="chart-title">Failures by Grade <span class="badge">Theory marks</span></div><div class="chart-sub">Failed ≥1 subject</div><div class="chart-wrap"><canvas id="chartGrade"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">Failures by Subject <span class="badge">Theory marks</span></div><div class="chart-sub">Subject-level failure count</div><div class="chart-wrap"><canvas id="chartSubject"></canvas></div></div>
    </div>
    <div class="charts-grid">
      <div class="chart-card"><div class="chart-title">Failure Count Distribution</div><div class="chart-sub">How many subjects did each failed student fail?</div><div class="chart-wrap"><canvas id="chartDist"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">Pass vs Fail by Grade</div><div class="chart-sub">Stacked breakdown per grade</div><div class="chart-wrap"><canvas id="chartStack"></canvas></div></div>
    </div>
    <div class="charts-grid" style="grid-template-columns:1fr">
      <div class="chart-card"><div class="chart-title">Average Marks by Subject (Failed Students)</div><div class="chart-sub">Average theory score of students who failed that subject, with pass mark reference</div><div class="chart-wrap tall"><canvas id="chartAvgMarks"></canvas></div></div>
    </div>
  </div>
  <!-- TABLE -->
  <div id="view-table" style="display:none">
    <div class="table-header-row">
      <div><div class="section-title">📋 Student Failure Details</div><div class="section-sub" id="tableSubtitle">Showing failed students</div></div>
    </div>
    <div class="table-wrap">
      <div class="table-scroll">
        <table>
          <thead><tr>
            <th>#</th><th>Student Name</th><th>Grade</th><th>Sec</th>
            <th>Orientation</th><th>Exam</th><th>Subjects Failed (Marks)</th>
            <th>Fail Count</th><th>All Theory Marks</th><th>Compare Exams</th>
          </tr></thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
      <div class="pagination" id="pagination"></div>
    </div>
  </div>
  <!-- TREND -->
  <div id="view-trend" style="display:none">
    <div class="charts-grid" style="grid-template-columns:1fr">
      <div class="chart-card"><div class="chart-title">Failure Trend Across All Exams</div><div class="chart-sub">Students who failed ≥1 subject per exam</div><div class="chart-wrap tall"><canvas id="chartTrend"></canvas></div></div>
    </div>
    <div class="charts-grid">
      <div class="chart-card"><div class="chart-title">Subject Failures by Exam</div><div class="chart-sub">Failures per subject across all exams</div><div class="chart-wrap tall"><canvas id="chartSubjectTrend"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">Grade-wise Pass Rate Trend</div><div class="chart-sub">Pass % per grade across exams</div><div class="chart-wrap tall"><canvas id="chartGradeTrend"></canvas></div></div>
    </div>
  </div>
</div>
<script>
const SUBJ_LABELS = {FL:'First Lang',SL:'Second Lang',TL:'Third Lang',Mat:'Maths',GS:'Gen. Science',SS:'Social Sci'};
const EXAM_KEYS = ['PRE_MID','MID_TERM','POST_MID','ANNUAL'];
const EXAM_LABELS = {PRE_MID:'Pre-Mid',MID_TERM:'Mid-Term',POST_MID:'Post-Mid',ANNUAL:'Annual'};
const CC = {red:'#ff4d6d',blue:'#4facfe',green:'#43e97b',purple:'#6c63ff',yellow:'#f9a825',orange:'#f97316',teal:'#06b6d4',pink:'#ec4899'};
let RAW_DATA=null, allRecords=[], filteredRecords=[], currentPage=1, charts={}, openRows=new Set();
const PAGE_SIZE=25;

function loadData(){
  RAW_DATA=DASHBOARD_DATA;
  allRecords=RAW_DATA.records;
  const selSec=document.getElementById('fSection'), selOri=document.getElementById('fOrientation');
  [...new Set(allRecords.map(r=>r.section))].sort().forEach(s=>{const o=document.createElement('option');o.value=s;o.textContent='Sec '+s;selSec.appendChild(o)});
  [...new Set(allRecords.map(r=>r.orientation))].sort().forEach(s=>{const o=document.createElement('option');o.value=s;o.textContent=s;selOri.appendChild(o)});
  applyFilters();
}

function applyFilters(){
  const exam=document.getElementById('fExam').value, grade=document.getElementById('fGrade').value,
    section=document.getElementById('fSection').value, orientation=document.getElementById('fOrientation').value,
    status=document.getElementById('fStatus').value, search=document.getElementById('fSearch').value.trim().toLowerCase();
  filteredRecords=allRecords.filter(r=>{
    if(exam!=='ALL'&&r.exam_key!==exam)return false;
    if(grade!=='ALL'&&String(r.grade)!==grade)return false;
    if(section!=='ALL'&&r.section!==section)return false;
    if(orientation!=='ALL'&&r.orientation!==orientation)return false;
    if(status==='FAILED'&&r.failed_count===0)return false;
    if(search&&!r.name.toLowerCase().includes(search)&&!r.admin_no.toLowerCase().includes(search))return false;
    return true;
  });
  currentPage=1; openRows=new Set();
  updateCards(); updateCharts(); updateTable();
}

function resetFilters(){
  document.getElementById('fExam').value='ANNUAL';
  document.getElementById('fGrade').value='ALL';
  document.getElementById('fSection').value='ALL';
  document.getElementById('fOrientation').value='ALL';
  document.getElementById('fStatus').value='FAILED';
  document.getElementById('fSearch').value='';
  applyFilters();
}

function updateCards(){
  const total=filteredRecords.length, failed=filteredRecords.filter(r=>r.failed_count>0).length,
    multi=filteredRecords.filter(r=>r.failed_count>=2).length, crit=filteredRecords.filter(r=>r.failed_count>=4).length;
  document.getElementById('cTotal').textContent=total;
  document.getElementById('cFailed').textContent=failed;
  document.getElementById('cPassRate').textContent=total>0?(((total-failed)/total)*100).toFixed(1)+'%':'—';
  document.getElementById('cMulti').textContent=multi;
  document.getElementById('cCritical').textContent=crit;
  document.getElementById('cFailedSub').textContent='of '+total+' records';
}

// --- CHARTS ---
function ax(){return{ticks:{color:'#8890b0',font:{size:11}},grid:{color:'#2e3148'},border:{color:'#2e3148'}}}
function chartOpts(y){return{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#8890b0',font:{size:11},boxWidth:24}},tooltip:{backgroundColor:'#1a1d27',borderColor:'#2e3148',borderWidth:1,titleColor:'#e8eaf6',bodyColor:'#8890b0',padding:10,cornerRadius:8}},scales:{x:ax(),y:{...ax()}}}}
function dc(id){if(charts[id]){charts[id].destroy();delete charts[id]}}

function updateCharts(){
  const fr=filteredRecords;
  const grades=[6,7,8,9];
  // Grade chart
  dc('g'); charts['g']=new Chart(document.getElementById('chartGrade'),{type:'bar',data:{labels:grades.map(g=>'Grade '+g),datasets:[{label:'Failed',data:grades.map(g=>fr.filter(r=>r.grade===g&&r.failed_count>0).length),backgroundColor:[CC.blue+'cc',CC.purple+'cc',CC.red+'cc',CC.yellow+'cc'],borderColor:[CC.blue,CC.purple,CC.red,CC.yellow],borderWidth:2,borderRadius:7}]},options:chartOpts('Students')});
  // Subject chart
  const SUBJS=['FL','SL','TL','Mat','GS','SS'], scols=[CC.blue,CC.purple,CC.teal,CC.red,CC.green,CC.yellow];
  dc('s'); charts['s']=new Chart(document.getElementById('chartSubject'),{type:'bar',data:{labels:SUBJS.map(s=>SUBJ_LABELS[s]),datasets:[{label:'Failures',data:SUBJS.map(s=>fr.reduce((a,r)=>a+(r.failed_subjects[s]!==undefined?1:0),0)),backgroundColor:scols.map(c=>c+'cc'),borderColor:scols,borderWidth:2,borderRadius:7}]},options:{...chartOpts('Failures'),indexAxis:'y'}});
  // Dist
  dc('d'); charts['d']=new Chart(document.getElementById('chartDist'),{type:'bar',data:{labels:['1 Subj','2 Subj','3 Subj','4 Subj','5 Subj','6 Subj'],datasets:[{label:'Students',data:[1,2,3,4,5,6].map(n=>fr.filter(r=>r.failed_count===n).length),backgroundColor:[CC.green+'cc',CC.blue+'cc',CC.yellow+'cc',CC.orange+'cc',CC.red+'cc',CC.pink+'cc'],borderColor:[CC.green,CC.blue,CC.yellow,CC.orange,CC.red,CC.pink],borderWidth:2,borderRadius:7}]},options:chartOpts('Students')});
  // Stack
  dc('st'); charts['st']=new Chart(document.getElementById('chartStack'),{type:'bar',data:{labels:grades.map(g=>'Grade '+g),datasets:[{label:'Passed',data:grades.map(g=>fr.filter(r=>r.grade===g&&r.failed_count===0).length),backgroundColor:CC.green+'cc',borderColor:CC.green,borderWidth:2,borderRadius:4},{label:'Failed',data:grades.map(g=>fr.filter(r=>r.grade===g&&r.failed_count>0).length),backgroundColor:CC.red+'cc',borderColor:CC.red,borderWidth:2,borderRadius:4}]},options:{...chartOpts('Students'),scales:{x:{stacked:true,...ax()},y:{stacked:true,...ax()}}}});
  // Avg marks
  const fld=fr.filter(r=>r.failed_count>0);
  const pm=fr.length>0?(fr[0].pass_mark||28):28;
  dc('a'); charts['a']=new Chart(document.getElementById('chartAvgMarks'),{type:'bar',data:{labels:SUBJS.map(s=>SUBJ_LABELS[s]),datasets:[{label:'Avg Marks (failed in subj)',data:SUBJS.map(s=>{const fi=fld.filter(r=>r.failed_subjects[s]!==undefined&&r.failed_subjects[s]!==null);return fi.length?parseFloat((fi.reduce((a,r)=>a+(r.failed_subjects[s]||0),0)/fi.length).toFixed(1)):0}),backgroundColor:CC.purple+'cc',borderColor:CC.purple,borderWidth:2,borderRadius:7},{label:'Pass Mark',data:SUBJS.map(()=>pm),type:'line',borderColor:CC.red,borderWidth:2,pointRadius:0,fill:false}]},options:chartOpts('Marks')});
  // Trend
  const grade=document.getElementById('fGrade').value, ori=document.getElementById('fOrientation').value, sec=document.getElementById('fSection').value;
  const tg=grade==='ALL'?[6,7,8]:[parseInt(grade)].filter(g=>g<=8);
  const tcols=[CC.blue,CC.purple,CC.green,CC.yellow];
  dc('tr'); charts['tr']=new Chart(document.getElementById('chartTrend'),{type:'line',data:{labels:EXAM_KEYS.map(k=>EXAM_LABELS[k]),datasets:tg.map((g,i)=>({label:'Grade '+g,data:EXAM_KEYS.map(ek=>{const rx=allRecords.filter(r=>r.exam_key===ek&&r.grade===g&&(ori==='ALL'||r.orientation===ori)&&(sec==='ALL'||r.section===sec));return rx.filter(r=>r.failed_count>0).length}),borderColor:tcols[i],backgroundColor:tcols[i]+'22',borderWidth:2.5,tension:.4,fill:true,pointRadius:5}))},options:chartOpts('Failed Students')});
  dc('sj'); charts['sj']=new Chart(document.getElementById('chartSubjectTrend'),{type:'line',data:{labels:EXAM_KEYS.map(k=>EXAM_LABELS[k]),datasets:SUBJS.map((s,i)=>({label:SUBJ_LABELS[s],data:EXAM_KEYS.map(ek=>allRecords.filter(r=>r.exam_key===ek&&(grade==='ALL'||String(r.grade)===grade)&&r.failed_subjects[s]!==undefined).length),borderColor:scols[i],backgroundColor:scols[i]+'22',borderWidth:2,tension:.4,fill:false,pointRadius:4}))},options:chartOpts('Failures')});
  dc('gt'); charts['gt']=new Chart(document.getElementById('chartGradeTrend'),{type:'line',data:{labels:EXAM_KEYS.map(k=>EXAM_LABELS[k]),datasets:tg.map((g,i)=>({label:'Grade '+g+' Pass%',data:EXAM_KEYS.map(ek=>{const rx=allRecords.filter(r=>r.exam_key===ek&&r.grade===g);return rx.length?parseFloat(((rx.filter(r=>r.failed_count===0).length/rx.length)*100).toFixed(1)):null}),borderColor:tcols[i],backgroundColor:tcols[i]+'22',borderWidth:2,tension:.4,fill:false,pointRadius:4}))},options:{...chartOpts('Pass %'),scales:{x:ax(),y:{...ax(),min:0,max:100}}}});
}

// --- TABLE ---
function updateTable(){
  const rows=filteredRecords.filter(r=>r.failed_count>0);
  const total=rows.length, pages=Math.ceil(total/PAGE_SIZE);
  if(currentPage>pages)currentPage=1;
  const start=(currentPage-1)*PAGE_SIZE, slice=rows.slice(start,start+PAGE_SIZE);
  document.getElementById('tableSubtitle').textContent=`Showing ${start+1}–${Math.min(start+PAGE_SIZE,total)} of ${total} failed records`;
  const tbody=document.getElementById('tableBody');
  if(!slice.length){tbody.innerHTML='<tr><td colspan="10" style="text-align:center;padding:50px;color:var(--text-muted)">✅ No failed students matching the current filter.</td></tr>';document.getElementById('pagination').innerHTML='';return;}
  tbody.innerHTML=slice.map((r,i)=>renderStudentRows(r, start+i+1)).join('');
  renderPagination(total,pages);
}

function renderStudentRows(r, idx){
  const fHtml=Object.entries(r.failed_subjects).map(([s,m])=>`<div class="subj-box${m===null?' ab':''}" title="${SUBJ_LABELS[s]||s}"><span class="sn">${s}</span><span class="sm">${m===null?'AB':m+'/'+r.total_theory}</span></div>`).join('');
  const allHtml=Object.entries(r.theory_marks).map(([s,m])=>`<span style="font-size:11px;color:${r.failed_subjects[s]!==undefined?'var(--red)':'var(--green)'}">${s}:${m===null?'AB':m}</span>`).join(' ');
  const gc={6:'blue',7:'purple',8:'yellow',9:'teal'}[r.grade]||'blue';
  const fc=r.failed_count>=4?'red':r.failed_count>=2?'yellow':'green';
  const hasHistory=r.exam_history&&Object.keys(r.exam_history).length>0;
  const isOpen=openRows.has(r.admin_no+'_'+idx);
  const compBtn=hasHistory?`<button class="toggle-btn${isOpen?' active':''}" onclick="toggleComp('${r.admin_no}',${idx})">${isOpen?'▲ Hide':'📊 Compare'}</button>`:`<span style="color:var(--text-muted);font-size:11px">—</span>`;
  const mainRow=`<tr class="data-row" id="row-${r.admin_no}-${idx}">
    <td style="color:var(--text-muted);font-size:11px">${idx}</td>
    <td><div class="td-name">${r.name}</div><div class="td-admin">${r.admin_no}</div></td>
    <td><span class="pill pill-${gc}">Gr ${r.grade}</span></td>
    <td><span class="pill pill-purple">${r.section}</span></td>
    <td style="font-size:11px;color:var(--text-muted)">${r.orientation}</td>
    <td><span class="pill pill-blue" style="font-size:10px">${r.exam_label}</span></td>
    <td><div style="display:flex;flex-wrap:wrap">${fHtml}</div></td>
    <td><span class="pill pill-${fc}">${r.failed_count}</span></td>
    <td style="white-space:nowrap">${allHtml}</td>
    <td>${compBtn}</td>
  </tr>`;
  const compRow=isOpen&&hasHistory?`<tr class="comp-row" id="comp-${r.admin_no}-${idx}"><td colspan="10" class="comp-panel"><div class="comp-inner">${buildCompPanel(r)}</div></td></tr>`:'';
  return mainRow+compRow;
}

function buildCompPanel(r){
  const exams=['PRE_MID','MID_TERM','POST_MID','ANNUAL'];
  const examData={'ANNUAL':{theory_marks:r.theory_marks,failed_subjects:r.failed_subjects,pass_mark:r.pass_mark,total_theory:r.total_theory,exam_label:r.exam_label},...r.exam_history};
  const subjects=Object.keys(r.theory_marks);
  // Header
  let html=`<div class="comp-title">📊 Cross-Exam Comparison for <strong style="color:var(--text)">${r.name}</strong> (${r.admin_no})</div>`;
  html+=`<table class="comp-table"><thead><tr><th class="subj-col">Subject</th>`;
  exams.forEach(ek=>{
    const d=examData[ek];
    if(!d)return;
    html+=`<th>${d.exam_label}<br><span style="font-weight:400;font-size:9px">(Pass: ${d.pass_mark}/${d.total_theory})</span></th>`;
  });
  html+=`</tr></thead><tbody>`;
  subjects.forEach(s=>{
    html+=`<tr><td class="subj-col">${SUBJ_LABELS[s]||s} <span style="opacity:.5">(${s})</span></td>`;
    exams.forEach(ek=>{
      const d=examData[ek];
      if(!d){html+=`<td class="c-na">—</td>`;return;}
      const mark=d.theory_marks[s];
      const pm=d.pass_mark;
      if(mark===null||mark===undefined){html+=`<td class="c-ab">AB</td>`;return;}
      const passed=mark>=pm;
      html+=`<td class="${passed?'c-pass':'c-fail'}">${mark}/${d.total_theory} ${passed?'✓':'✗'}</td>`;
    });
    html+=`</tr>`;
  });
  // Summary row
  html+=`<tr style="border-top:2px solid var(--border)"><td class="subj-col" style="font-weight:700;color:var(--text)">Subjects Failed</td>`;
  exams.forEach(ek=>{
    const d=examData[ek];
    if(!d){html+=`<td class="c-na">—</td>`;return;}
    const fc=d.failed_count;
    html+=`<td class="${fc===0?'c-pass':fc>=4?'c-fail':'c-ab'}" style="font-weight:700">${fc} fail${fc!==1?'s':''}</td>`;
  });
  html+=`</tr></tbody></table>`;
  return html;
}

function toggleComp(adminNo, idx){
  const key=adminNo+'_'+idx;
  if(openRows.has(key))openRows.delete(key); else openRows.add(key);
  updateTable();
  // Scroll to keep row in view
  setTimeout(()=>{const el=document.getElementById('row-'+adminNo+'-'+idx);if(el)el.scrollIntoView({behavior:'smooth',block:'nearest'})},50);
}

function renderPagination(total,pages){
  if(!total){document.getElementById('pagination').innerHTML='';return;}
  let h=`<span class="page-info">Page ${currentPage} of ${pages} (${total} records)</span>`;
  h+=`<button class="page-btn" onclick="goPage(${currentPage-1})" ${currentPage===1?'disabled':''}>◀</button>`;
  const rng=pages<=7?Array.from({length:pages},(_,i)=>i+1):currentPage<=4?Array.from({length:7},(_,i)=>i+1):currentPage>=pages-3?Array.from({length:7},(_,i)=>pages-6+i):Array.from({length:7},(_,i)=>currentPage-3+i);
  rng.forEach(p=>h+=`<button class="page-btn${p===currentPage?' active':''}" onclick="goPage(${p})">${p}</button>`);
  h+=`<button class="page-btn" onclick="goPage(${currentPage+1})" ${currentPage===pages?'disabled':''}>▶</button>`;
  document.getElementById('pagination').innerHTML=h;
}

function goPage(p){
  const rows=filteredRecords.filter(r=>r.failed_count>0);
  const pages=Math.ceil(rows.length/PAGE_SIZE);
  if(p<1||p>pages)return;
  currentPage=p; openRows=new Set(); updateTable();
}

function switchTab(t){
  ['charts','table','trend'].forEach(k=>{
    document.getElementById('view-'+k).style.display=k===t?'':'none';
    document.getElementById('tab-'+k).classList.toggle('active',k===t);
  });
}

function downloadExcel(){
  const exam=document.getElementById('fExam').value, grade=document.getElementById('fGrade').value;
  const rows=filteredRecords.filter(r=>r.failed_count>0);
  if(!rows.length){alert('No failed students to export.');return;}
  const wb=XLSX.utils.book_new();
  const d1=rows.map(r=>({'Student Name':r.name,'Admin No':r.admin_no,'Grade':r.grade,'Section':r.section,'Orientation':r.orientation,'Exam':r.exam_label,'Subjects Failed':Object.keys(r.failed_subjects).join(', '),'Failed Count':r.failed_count,'FL Theory':r.theory_marks['FL']??'AB','SL Theory':r.theory_marks['SL']??'AB','TL Theory':r.theory_marks['TL']??'AB','Mat Theory':r.theory_marks['Mat']??'AB','GS Theory':r.theory_marks['GS']??'AB','SS Theory':r.theory_marks['SS']??'AB','Pass Mark':r.pass_mark,'Max Theory':r.total_theory}));
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d1),'Failure Details');
  // Cross-exam comparison sheet for annual failures
  const annualFailed=rows.filter(r=>r.exam_key==='ANNUAL');
  if(annualFailed.length){
    const d2=[];
    annualFailed.forEach(r=>{
      const EXAMS=['PRE_MID','MID_TERM','POST_MID','ANNUAL'];
      const allExamData={...r.exam_history,ANNUAL:{theory_marks:r.theory_marks,failed_subjects:r.failed_subjects,pass_mark:r.pass_mark,total_theory:r.total_theory,exam_label:r.exam_label}};
      Object.keys(r.theory_marks).forEach(s=>{
        const row={'Student Name':r.name,'Admin No':r.admin_no,'Grade':r.grade,'Section':r.section,'Subject':SUBJ_LABELS[s]||s};
        EXAMS.forEach(ek=>{const d=allExamData[ek];if(!d){row[d?.exam_label||ek]='N/A';return;}const m=d.theory_marks[s];row[d.exam_label+' ('+d.total_theory+')']=m===null?'AB':m;row[d.exam_label+' P/F']=m===null?'AB':m>=d.pass_mark?'PASS':'FAIL';});
        d2.push(row);
      });
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d2),'Cross-Exam Comparison');
  }
  // Grade summary
  const d3=[6,7,8,9].flatMap(g=>{const gr=rows.filter(r=>r.grade===g);if(!gr.length)return[];const tot=allRecords.filter(r=>r.grade===g&&(exam==='ALL'||r.exam_key===exam)).length;return[{'Grade':g,'Total Students':tot,'Failed':gr.length,'Pass%':tot?((1-gr.length/tot)*100).toFixed(1)+'%':'—'}]});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d3),'Grade Summary');
  XLSX.writeFile(wb,'SCTS_Failure_'+exam+'_Grade'+grade+'.xlsx');
}

loadData();
</script>
</body></html>"""

if __name__ == '__main__':
    data = build_data()
    json_path = os.path.join(BASE_DIR, 'dashboard_data.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    print(f'JSON saved: {json_path}')
    inline_js = f'const DASHBOARD_DATA = {json.dumps(data, ensure_ascii=False)};'
    html = HTML_TEMPLATE.replace('loadData();', f'{inline_js}\nloadData();')
    out_path = os.path.join(BASE_DIR, 'dashboard.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Dashboard built: {out_path}')
    print('Open dashboard.html in any browser - no server needed!')
