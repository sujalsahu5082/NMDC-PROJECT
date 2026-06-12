let DATA = [], PAST = [], COLS = [], ROWS = [], MAP = {}, CH = {}, DCH = {};
let DRILL_DATA = [], DRILL_ALL = [], DRILL_TITLE = '';
let SORT_COL = null, SORT_DIR = 1, CUR_PAGE = 1, PAGE_SIZE = 25;
let FILTERED_CACHE = [];
let RPT_FILTERED = [];
let FS_TYPE = null, FS_CH = null, FS_TABLE_DATA = [], FS_TABLE_KEY = 'Item';

const TODAY = new Date();
TODAY.setHours(0, 0, 0, 0);

if (typeof Chart === 'undefined') { console.error('Chart.js not loaded!'); alert('Chart.js library not loaded. Please check your internet connection and refresh the page.'); }
if (typeof XLSX === 'undefined') { console.error('SheetJS not loaded!'); alert('Excel library not loaded. Please check your internet connection and refresh the page.'); }

// ── SINGLE GLOBAL FILTER STATE ──
let GF = { dep: [], dept: [], sub: [], skill: [], grade: [], desig: [], status: '', ry: [], rfy: [], from: '', to: '' };

// ── FS LOCAL STATE ──
let FS_LOCAL = { timePeriod: '', clickedType: null, clickedValues: [] };

// ── SELECTION STATE ──
let SELECTED_IDS = new Set();
let DRILL_SELECTED_IDS = new Set();

let _lastSkillMap = {}, _lastDeptMap = {}, _lastDepMap = {}, _lastGradeMap = {}, _lastFySorted = [], _lastDgMap = {}, _lastSkillGap = [], _lastShortfall = [];

const ORG = {
    production: ['Mining', 'Services (Mech.)', 'Services (Elect.)', 'Plant (Mech.) incl. L&D', 'Plant (Elect.)', 'Geology & QC', 'Chemical'],
    nonproduction: ['Civil', 'Materials', 'T&S and Environment', 'Finance', 'Human Resource, RB & CSR', 'Industrial Engineering', 'M&S', 'C & IT', 'ED Sectt.', 'CGM (P) Sectt.', 'Vigilance', 'Contracts', 'SP-III / Works Sectt.'],
    others: ['School', 'Hospital']
};

function classifyCategory(dept) {
    if (!dept || dept === 'N/A') return 'Unclassified';
    const d = dept.trim().toLowerCase();
    if (ORG.production.some(p => d.includes(p.toLowerCase()) || p.toLowerCase().includes(d))) return 'Production';
    if (ORG.nonproduction.some(p => d.includes(p.toLowerCase()) || p.toLowerCase().includes(d))) return 'Non-Production';
    return 'Others';
}

function empKey(e) { return (e.empid || '') + '||' + e.name + '||' + (e.dob instanceof Date ? e.dob.getTime() : e.dob); }

// ===================== HELPERS =====================
function getFY(d) { const y = d.getFullYear(), m = d.getMonth(); return m >= 3 ? `FY ${y}-${String(y+1).slice(2)}` : `FY ${y-1}-${String(y).slice(2)}`; }
function getFYYear(fy) { const m = fy.match(/FY (\d{4})/); return m ? +m[1] : 0; }
function fmtDate(d) { return d ? d.toLocaleDateString('en-IN', {day:'2-digit',month:'short',year:'numeric'}) : '—'; }
function fmtYMD(d) { return d.toISOString().split('T')[0]; }
function esc(s) { return (s||'').toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function calcRetDate(dob) { const b = new Date(dob); b.setFullYear(b.getFullYear()+60); b.setHours(0,0,0,0); return b.getDate()===1 ? new Date(b.getFullYear(),b.getMonth(),0) : new Date(b.getFullYear(),b.getMonth()+1,0); }
function get60th(dob) { const b = new Date(dob); b.setFullYear(b.getFullYear()+60); b.setHours(0,0,0,0); return b; }
function addMonths(n) { const d = new Date(TODAY); d.setMonth(d.getMonth()+n); return d; }
function addYears(n) { const d = new Date(TODAY); d.setFullYear(d.getFullYear()+n); return d; }
function uniq(arr) { return [...new Set(arr.filter(v=>v&&v!=='N/A'))].sort(); }
function fmtN(n) { return Number(n).toLocaleString('en-IN'); }
function getStatus(e) { if (e.rd < TODAY) return 'retired'; if (e.days <= 90) return 'soon'; if (e.days <= 365) return 'year'; return 'active'; }
function calcAge(dob) { return Math.floor((TODAY - dob) / 31556952000); }
function riskLvl(p) { return p >= 40 ? 'critical' : p >= 25 ? 'high' : p >= 10 ? 'medium' : 'low'; }
function riskBadge(lv) { const icons={critical:'⚠',high:'🔴',medium:'🟡',low:'🟢'}; return `<span class="risk-${lv}">${icons[lv]} ${lv[0].toUpperCase()+lv.slice(1)}</span>`; }
function inH(e, h) { return e.rd >= TODAY && e.rd <= h; }
function retCalcExpl(dob, b60, rd) { return `<strong>60th Birthday:</strong> ${fmtDate(b60)}<br><strong>Rule:</strong> ${b60.getDate()===1?'1st → last day of prev month':'Not 1st → last day of same month'}<br><strong>Retirement Date:</strong> ${fmtDate(rd)}`; }
function statusBadge(e) { const s=getStatus(e); if(s==='retired')return`<span class="status-retired">Retired</span>`; if(s==='soon')return`<span class="status-soon">Soon</span>`; if(s==='year')return`<span class="status-year">This Year</span>`; return`<span class="status-active">Active</span>`; }
function showLoading(msg) { document.getElementById('loading-text').textContent=msg||'Processing…'; document.getElementById('loading-ov').classList.add('show'); }
function hideLoading() { document.getElementById('loading-ov').classList.remove('show'); }

// ===================== TOAST =====================
let _toastTimer = null;
function showToast(msg, icon='fa-check-circle') {
    const t = document.getElementById('toast');
    t.querySelector('i').className = 'fas '+icon;
    document.getElementById('toast-msg').textContent = msg;
    t.classList.add('show');
    if (_toastTimer) clearTimeout(_toastTimer);
    _toastTimer = setTimeout(()=>t.classList.remove('show'), 3000);
}

// ===================== THEME =====================
let DARK = false;
function applyTheme(dark) {
    DARK = dark;
    document.body.classList.toggle('dark', DARK);
    const btn = document.getElementById('theme-btn');
    if (btn) btn.innerHTML = DARK ? '<i class="fas fa-sun"></i> Light Mode' : '<i class="fas fa-moon"></i> Dark Mode';
    if (DATA.length || PAST.length) setTimeout(drawAllCharts, 50);
}
function toggleTheme() { applyTheme(!DARK); }

// ── postMessage bridge: parent NMDC dashboard → retirement iframe ──────────────
window.addEventListener('message', function(e) {
    if (!e.data || typeof e.data !== 'object') return;
    const { type, payload } = e.data;

    if (type === 'setTheme') {
        // payload: { dark: true|false }
        applyTheme(!!payload.dark);
    }

    if (type === 'clearData') {
        // Reset all data and return to upload landing screen
        DATA = []; PAST = []; COLS = []; ROWS = []; MAP = {};
        Object.values(CH).forEach(c => { try { c.destroy(); } catch (_) {} });
        Object.values(DCH).forEach(c => { try { c.destroy(); } catch (_) {} });
        CH = {}; DCH = {};
        GF = { dep: [], dept: [], sub: [], skill: [], grade: [], desig: [], status: '', ry: [], rfy: [], from: '', to: '' };
        FILTERED_CACHE = []; RPT_FILTERED = [];
        SELECTED_IDS.clear(); DRILL_SELECTED_IDS.clear();
        // Hide dashboards, show landing
        const land = document.getElementById('land');
        const empDash = document.getElementById('emp-dash');
        const fileInfoInline = document.getElementById('file-info-inline');
        if (land) land.style.display = '';
        if (empDash) empDash.style.display = 'none';
        if (fileInfoInline) fileInfoInline.style.display = 'none';
        // Reset employee count badge
        const ecnt = document.getElementById('ecnt');
        if (ecnt) ecnt.textContent = '0';
        showToast('Data cleared', 'fa-trash');
    }
});

// ===================== MULTI-SELECT =====================
const MS_DATA = {};
const MS_SUFFIXES = ['-emp', '-charts', '-report', '-fs'];

function initMS(id, values, onChg) {
    MS_DATA[id] = {values, selected: [], onChg};
    MS_SUFFIXES.forEach(suffix => {
        const opts = document.getElementById(id + suffix + '-opts');
        if (opts) {
            opts.innerHTML = values.map(v=>`<label class="ms-opt"><input type="checkbox" value="${esc(v)}" onchange="onMSChange('${id}', '${suffix}')"> ${esc(v)}</label>`).join('');
            updateMSTrigger(id, suffix);
        }
    });
}

function onMSChange(id, suffix) {
    const currentOpts = document.getElementById(id + suffix + '-opts');
    if (!currentOpts) return;
    const chks = currentOpts.querySelectorAll('input[type=checkbox]');
    const selected = [...chks].filter(c=>c.checked).map(c=>c.value);
    MS_DATA[id].selected = selected;
    
    // Sync other suffixes
    MS_SUFFIXES.forEach(s => {
        if (s !== suffix) {
            const otherOpts = document.getElementById(id + s + '-opts');
            if (otherOpts) {
                const otherChks = otherOpts.querySelectorAll('input[type=checkbox]');
                otherChks.forEach(c => {
                    c.checked = selected.includes(c.value);
                });
            }
        }
    });
    
    // Update trigger labels for all suffixes
    MS_SUFFIXES.forEach(s => {
        updateMSTrigger(id, s);
    });
    
    if (MS_DATA[id].onChg) MS_DATA[id].onChg();
}

function updateMSTrigger(id, suffix) {
    const t = document.getElementById(id + suffix + '-trigger');
    if (!t) return;
    const s = MS_DATA[id]?.selected || [];
    t.textContent = s.length===0 ? (id.includes('dep')&&!id.includes('dept') ? 'All Deposits' : id.includes('ry') ? 'All Years' : id.includes('rfy') ? 'All FY' : 'All') : s.length===1 ? s[0] : `${s.length} selected`;
}

function toggleMS(id) {
    document.querySelectorAll('.ms-panel.open').forEach(p => {
        if (p.id !== id) { p.classList.remove('open'); const tr=document.getElementById(p.id+'-trigger'); if(tr)tr.classList.remove('open'); }
    });
    const p=document.getElementById(id), t=document.getElementById(id+'-trigger');
    if(p){p.classList.toggle('open');if(t)t.classList.toggle('open');}
}

function filterMSOpts(inp) {
    const q = inp.value.toLowerCase();
    inp.closest('.ms-panel').querySelectorAll('.ms-opt').forEach(o=>{o.style.display=o.textContent.toLowerCase().includes(q)?'':'none';});
}

function selectAllMS(idWithSuffix) {
    const suffix = MS_SUFFIXES.find(s => idWithSuffix.endsWith(s));
    if (!suffix) return;
    const id = idWithSuffix.slice(0, -suffix.length);
    
    document.querySelectorAll(`#${idWithSuffix}-opts input[type=checkbox]`).forEach(c=>{if(c.closest('.ms-opt').style.display!=='none')c.checked=true;});
    onMSChange(id, suffix);
}

function clearMS(idWithSuffix) {
    const suffix = MS_SUFFIXES.find(s => idWithSuffix.endsWith(s));
    if (!suffix) return;
    const id = idWithSuffix.slice(0, -suffix.length);
    
    document.querySelectorAll(`#${idWithSuffix}-opts input[type=checkbox]`).forEach(c=>c.checked=false);
    onMSChange(id, suffix);
}

function getMSVals(id) { return MS_DATA[id]?.selected || []; }

document.addEventListener('click', e => {
    if (!e.target.closest('.ms-wrap') && !e.target.closest('.ms-panel')) {
        document.querySelectorAll('.ms-panel.open').forEach(p => {
            p.classList.remove('open');
            const t = document.getElementById(p.id+'-trigger');
            if (t) t.classList.remove('open');
        });
    }
    const expDd = document.getElementById('exp-dd');
    if (expDd && !e.target.closest('.exp-wrap')) expDd.classList.remove('open');
});

// ===================== FILE UPLOAD =====================
const FI = document.getElementById('fi-inline');
const inlineUploadBtn = document.getElementById('inline-upload-btn');
if (inlineUploadBtn) inlineUploadBtn.addEventListener('click', ()=>FI.click());
FI.addEventListener('click', e => { e.target.value = null; });
FI.addEventListener('change', e=>{if(e.target.files[0])readF(e.target.files[0]);});

const dz = document.getElementById('dz');
if (dz) {
    dz.addEventListener('click', ()=>FI.click());
    dz.addEventListener('dragover', e=>{e.preventDefault();dz.classList.add('drag');});
    dz.addEventListener('dragleave', ()=>dz.classList.remove('drag'));
    dz.addEventListener('drop', e=>{e.preventDefault();dz.classList.remove('drag');if(e.dataTransfer.files[0])readF(e.dataTransfer.files[0]);});
}

function readF(file) {
    showLoading('Reading file…');
    setTimeout(()=>{
        const fr = new FileReader();
        fr.onload = e => {
            try {
                const wb = XLSX.read(e.target.result,{type:'binary',cellDates:true});
                const ws = wb.Sheets[wb.SheetNames[0]];
                ROWS = XLSX.utils.sheet_to_json(ws,{raw:false,dateNF:'DD-MM-YYYY'});
                if (!ROWS.length){hideLoading();alert('No data found.');return;}
                COLS = Object.keys(ROWS[0]);
                document.getElementById('fcn-inline').textContent = file.name;
                document.getElementById('file-info-inline').style.display = 'inline-flex';
                // Auto-detect columns — no modal needed
                autoMapAndParse();
                hideLoading();
                showToast('Data loaded successfully!','fa-check-circle');
            } catch(err){hideLoading();alert('Error: '+err.message);}
        };
        fr.readAsBinaryString(file);
    }, 50);
}

// ===================== AUTO COLUMN DETECTION =====================
// Extended hints that match exact NMDC Excel column names
const FIELDS = [
    {k:'name',   lbl:'Employee Name',         req:false, h:['n a m e','name','employee name','emp name','full name','emp_name']},
    {k:'empid',  lbl:'Employee ID',            req:false, h:['uec no','uec no.','sap uec','empid','emp id','employee id','staff id','personnel']},
    {k:'dob',    lbl:'Date of Birth',          req:true,  h:['dob','date of birth','birthdate','birth date','d.o.b']},
    {k:'dor',    lbl:'Date of Retirement',     req:false, h:['dor','date of retirement','retirement date','ret. date','ret date']},
    {k:'deposit',lbl:'Deposit (11B/11C/14)',   req:false, h:['dc',' dc','deposit','mine','pit','unit']},
    {k:'dept',   lbl:'Department',             req:false, h:['department','dept','division']},
    {k:'subdept',lbl:'Sub Department',         req:false, h:['section','sub department','subdept','sub-dept','sub dept']},
    {k:'grade',  lbl:'Grade',                  req:false, h:['grade','pay grade','level','pg grade']},
    {k:'desig',  lbl:'Designation',            req:false, h:['designation','post','position','title','desig']},
    {k:'skill',  lbl:'Skill / Trade',          req:false, h:['original skill','skill','trade','specialisation','competency']},
    {k:'gender', lbl:'Gender',                 req:false, h:['gender','sex']},
];

function bestM(hints) {
    for (const c of COLS) {
        const cl = c.trim().toLowerCase();
        for (const h of hints) {
            if (cl === h || cl.includes(h) || h.includes(cl)) return c;
        }
    }
    return '';
}

function autoMapAndParse() {
    MAP = {};
    FIELDS.forEach(f => {
        const v = bestM(f.h);
        if (v) MAP[f.k] = v;
    });
    if (!MAP.dob && !MAP.dor) {
        alert('Could not find DOB or Retirement Date column. Please check your file.');
        return;
    }
    showLoading('Processing retirement data…');
    setTimeout(()=>{ parse(); hideLoading(); }, 50);
}

// Stubs — mapping modal is disabled; auto-detection used instead
function openModal()  { /* no-op */ }
function closeModal() { document.getElementById('ov').style.display='none'; }
function procMap()    { closeModal(); autoMapAndParse(); }

function parseDateVal(raw) {
    if (!raw) return null;
    if (raw instanceof Date) {
        const d = new Date(raw);
        if (!isNaN(d) && d.getFullYear() > 1900) { d.setHours(0,0,0,0); return d; }
        return null;
    }
    const str = raw.toString().trim();
    // ISO: 2031-03-17 or 2031-03-17 00:00:00
    let m = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m) { const d = new Date(+m[1],+m[2]-1,+m[3]); if (!isNaN(d)) { d.setHours(0,0,0,0); return d; } }
    // DD-MM-YYYY or DD/MM/YYYY
    m = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (m) { const d = new Date(+m[3],+m[2]-1,+m[1]); if (!isNaN(d)) { d.setHours(0,0,0,0); return d; } }
    let d = new Date(str);
    if (!isNaN(d) && d.getFullYear() > 1900) { d.setHours(0,0,0,0); return d; }
    // Excel serial
    const sr = parseFloat(str);
    if (!isNaN(sr) && sr > 1000 && sr < 100000) {
        d = new Date(new Date(1899,11,30).getTime() + sr*86400000);
        d.setHours(0,0,0,0); return d;
    }
    return null;
}

// Kept as alias for legacy compatibility
function parseDOB(raw) { return parseDateVal(raw); }

function cleanDeposit(raw) {
    if (!raw) return 'N/A';
    // Strip "Dep-" prefix e.g. "Dep-11C" → "11C"
    return raw.toString().trim().replace(/^Dep[-\s]*/i, '').toUpperCase() || 'N/A';
}

function parse() {
    DATA=[];PAST=[];let skip=0;
    ROWS.forEach(row=>{
        // --- DOB ---
        const dob = parseDateVal(MAP.dob ? row[MAP.dob] : null);
        // --- Retirement Date: use DOR column if present, else calculate from DOB ---
        let rd = null;
        if (MAP.dor) rd = parseDateVal(row[MAP.dor]);
        if (!rd && dob) rd = calcRetDate(dob);
        if (!rd) { skip++; return; }
        rd.setHours(0,0,0,0);

        const b60 = dob ? get60th(dob) : rd;
        const days = Math.round((rd - TODAY) / 86400000);
        const st = days<=90 ? 'u' : days<=365 ? 's' : 'ok';

        const gRaw = MAP.gender ? (row[MAP.gender]||'').toString().trim() : '';
        const gl = gRaw.toLowerCase();
        let gender = '?';
        if (['m','male'].includes(gl) || gl.startsWith('m')) gender = 'M';
        else if (['f','female'].includes(gl) || gl.startsWith('f')) gender = 'F';
        else if (gRaw) gender = 'O';

        const dept = MAP.dept ? (row[MAP.dept]||'N/A').toString().trim() : 'N/A';
        const rawDeposit = MAP.deposit ? row[MAP.deposit] : '';
        const deposit = cleanDeposit(rawDeposit);

        const obj = {
            name:    MAP.name    ? (row[MAP.name]||'—').toString().trim() : '—',
            dob:     dob || rd,
            b60,
            deposit,
            category: classifyCategory(dept),
            dept,
            subdept: MAP.subdept ? (row[MAP.subdept]||'N/A').toString().trim() : 'N/A',
            grade:   MAP.grade   ? (row[MAP.grade]||'N/A').toString().trim() : 'N/A',
            desig:   MAP.desig   ? (row[MAP.desig]||'N/A').toString().trim() : 'N/A',
            skill:   MAP.skill   ? (row[MAP.skill]||'N/A').toString().trim() : 'N/A',
            gender,
            empid:   MAP.empid   ? (row[MAP.empid]||'').toString().trim() : '',
            rd,
            fy:      getFY(rd),
            fyYear:  getFYYear(getFY(rd)),
            rdMonth: rd.toLocaleDateString('en-IN',{month:'short',year:'numeric'}),
            rdMS:    `${rd.getFullYear()}-${String(rd.getMonth()+1).padStart(2,'0')}`,
            age:     dob ? calcAge(dob) : 0,
            days,
            st
        };
        if (rd < TODAY) PAST.push(obj); else DATA.push(obj);
    });
    DATA.sort((a,b)=>a.rd-b.rd);
    PAST.sort((a,b)=>b.rd-a.rd);
    if (skip) console.warn(`Skipped ${skip} rows (no valid date found)`);
    boot();
}

function boot() {
    document.getElementById('land').style.display = 'none';
    document.getElementById('emp-dash').style.display = 'block';
    document.getElementById('ch-ph').style.display = 'none';
    document.getElementById('ch-body').style.display = 'block';
    document.getElementById('rpt-ph').style.display = 'none';
    document.getElementById('rpt-body').style.display = 'block';
    document.getElementById('ecnt').textContent = DATA.length + PAST.length;

    CUR_PAGE=1;SORT_COL=null;SORT_DIR=1;
    GF={dep:[],dept:[],sub:[],skill:[],grade:[],desig:[],status:'',ry:[],rfy:[],from:'',to:''};
    SELECTED_IDS.clear();updateSelectionToolbar();

    const all=[...DATA,...PAST];
    const depts=uniq(all.map(e=>e.dept)),skills=uniq(all.map(e=>e.skill));
    const grades=['RS 10','RS 09','RS 08','RS 07','RS 06','RS 05','RS 04','RS 03','RS 02','RS 01'];
    const subs=uniq(all.map(e=>e.subdept)),desigs=uniq(all.map(e=>e.desig));
    const rys=uniq(all.map(e=>String(e.rd.getFullYear()))),rfys=uniq(all.map(e=>e.fy));
    const deposits=uniq(all.map(e=>e.deposit));
    const priorityDeps=['14','11C','11B'];
    const sortedDeposits=[...priorityDeps.filter(d=>deposits.includes(d)),...deposits.filter(d=>!priorityDeps.includes(d))];

    initMS('gf-dep', sortedDeposits, applyGF);
    initMS('gf-dept', depts, applyGF);
    initMS('gf-sub', subs, applyGF);
    initMS('gf-skill', skills, applyGF);
    initMS('gf-grade', grades, applyGF);
    initMS('gf-desig', desigs, applyGF);
    initMS('gf-ry', rys, applyGF);
    initMS('gf-rfy', rfys, applyGF);

    // Clear standard inputs across all suffixes
    MS_SUFFIXES.forEach(s => {
        const st = document.getElementById('gf-status' + s);
        const fr = document.getElementById('gf-from' + s);
        const to = document.getElementById('gf-to' + s);
        if (st) st.value = '';
        if (fr) fr.value = '';
        if (to) to.value = '';
    });

    document.getElementById('lf-d').innerHTML='<option value="">All</option>'+depts.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
    document.getElementById('rpt-date').textContent=TODAY.toLocaleDateString('en-IN',{day:'2-digit',month:'long',year:'numeric'});

    RPT_FILTERED=[...all];
    updateKPIs();
    renderTable();
    drawAllCharts();
    renderReport();
}

// ===================== GLOBAL FILTER =====================
function applyGFFrom(tabName) {
    const suffix = '-' + tabName;
    const statusEl = document.getElementById('gf-status' + suffix);
    const fromEl = document.getElementById('gf-from' + suffix);
    const toEl = document.getElementById('gf-to' + suffix);
    
    if (statusEl) GF.status = statusEl.value;
    if (fromEl) GF.from = fromEl.value;
    if (toEl) GF.to = toEl.value;
    
    // Sync other suffixes
    MS_SUFFIXES.forEach(s => {
        if (s !== suffix) {
            const st = document.getElementById('gf-status' + s);
            const fr = document.getElementById('gf-from' + s);
            const to = document.getElementById('gf-to' + s);
            if (st && statusEl) st.value = statusEl.value;
            if (fr && fromEl) fr.value = fromEl.value;
            if (to && toEl) to.value = toEl.value;
        }
    });
    
    applyGF();
}

function applyGF() {
    // Determine which suffix to read status/from/to from
    const fsOpen = document.getElementById('fs-ov') && document.getElementById('fs-ov').classList.contains('show');
    let suffix;
    if (fsOpen) {
        suffix = '-fs';
    } else {
        const activeTabBtn = document.querySelector('.tab.on');
        let tabName = 'emp';
        if (activeTabBtn) {
            const p = activeTabBtn.dataset.p;
            if (p === 'employees') tabName = 'emp';
            else if (p === 'charts') tabName = 'charts';
            else if (p === 'report') tabName = 'report';
        }
        suffix = '-' + tabName;
    }

    const statusEl = document.getElementById('gf-status' + suffix);
    const fromEl = document.getElementById('gf-from' + suffix);
    const toEl = document.getElementById('gf-to' + suffix);

    if (statusEl) GF.status = statusEl.value;
    if (fromEl) GF.from = fromEl.value;
    if (toEl) GF.to = toEl.value;

    // Sync all standard inputs to match GF state
    MS_SUFFIXES.forEach(s => {
        const st = document.getElementById('gf-status' + s);
        const fr = document.getElementById('gf-from' + s);
        const to = document.getElementById('gf-to' + s);
        if (st) st.value = GF.status || '';
        if (fr) fr.value = GF.from || '';
        if (to) to.value = GF.to || '';
    });

    GF.dep   = getMSVals('gf-dep');
    GF.dept  = getMSVals('gf-dept');
    GF.sub   = getMSVals('gf-sub');
    GF.skill = getMSVals('gf-skill');
    GF.grade = getMSVals('gf-grade');
    GF.desig = getMSVals('gf-desig');
    GF.ry   = getMSVals('gf-ry');
    GF.rfy  = getMSVals('gf-rfy');

    const pool = applyGFToPool([...DATA,...PAST]);
    RPT_FILTERED = pool;

    renderGFChips();
    updateGFActiveBadge();
    updateKPIs();
    renderTable();
    drawAllCharts();
    renderReport();
    updateInsightStrip();

    if (fsOpen) {
        FS_LOCAL.clickedType = null;
        FS_LOCAL.clickedValues = [];
        refreshFS();
    }
}

function clearAllGF() {
    GF={dep:[],dept:[],sub:[],skill:[],grade:[],desig:[],status:'',ry:[],rfy:[],from:'',to:''};
    ['gf-dep','gf-dept','gf-sub','gf-skill','gf-grade','gf-desig','gf-ry','gf-rfy'].forEach(id=>{
        if(MS_DATA[id]){
            MS_DATA[id].selected=[];
            MS_SUFFIXES.forEach(s => {
                const opts = document.getElementById(id + s + '-opts');
                if (opts) {
                    opts.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = false);
                }
                updateMSTrigger(id, s);
            });
        }
    });
    
    MS_SUFFIXES.forEach(s => {
        const st = document.getElementById('gf-status' + s);
        const fr = document.getElementById('gf-from' + s);
        const to = document.getElementById('gf-to' + s);
        if (st) st.value = '';
        if (fr) fr.value = '';
        if (to) to.value = '';
    });

    FS_LOCAL.clickedType=null;FS_LOCAL.clickedValues=[];FS_LOCAL.timePeriod='';
    updateFSTimeBtns();

    const pool=[...DATA,...PAST];
    RPT_FILTERED=pool;
    renderGFChips();
    updateGFActiveBadge();
    updateKPIs();
    renderTable();
    drawAllCharts();
    renderReport();
    document.getElementById('emp-insight-strip').style.display='none';
    showToast('All filters cleared','fa-rotate-left');
    if(document.getElementById('fs-ov') && document.getElementById('fs-ov').classList.contains('show')) {
        syncFSSidebarFromGF();
        refreshFS();
    }
}

function updateGFActiveBadge() {
    const badge = document.getElementById('gf-active-badge');
    if (!badge) return;
    const count = GF.dep.length + GF.dept.length + GF.sub.length + GF.skill.length + GF.grade.length + GF.desig.length + GF.ry.length + GF.rfy.length + (GF.status?1:0) + (GF.from?1:0) + (GF.to?1:0);
    if (count > 0) {
        badge.textContent = count + ' active';
        badge.style.display = 'inline-flex';
    } else {
        badge.style.display = 'none';
    }
}

function toggleGF(name) {
    const body = document.getElementById('gf-body-' + name);
    const icon = document.getElementById('gf-icon-' + name);
    if (!body) return;
    const isOpen = body.style.display !== 'none';
    body.style.display = isOpen ? 'none' : 'block';
    if (icon) icon.style.transform = isOpen ? '' : 'rotate(180deg)';
}

function renderGFChips() {
    let chips = '';
    const addChips = (vals, label, key) => vals.forEach(v=>{chips+=`<span class="gf-chip">${label}: <strong>${esc(v)}</strong><span class="gf-chip-x" onclick="removeGFVal('${key}','${esc(v)}')">×</span></span>`;});
    addChips(GF.dep,'Deposit','dep');addChips(GF.dept,'Dept','dept');addChips(GF.sub,'Sub','sub');
    addChips(GF.skill,'Skill','skill');addChips(GF.grade,'Grade','grade');addChips(GF.desig,'Desig','desig');
    addChips(GF.ry,'Year','ry');addChips(GF.rfy,'FY','rfy');
    if(GF.status)chips+=`<span class="gf-chip">Status: <strong>${GF.status}</strong><span class="gf-chip-x" onclick="removeGFVal('status','')">×</span></span>`;
    if(GF.from)chips+=`<span class="gf-chip">From: <strong>${GF.from}</strong><span class="gf-chip-x" onclick="removeGFVal('from','')">×</span></span>`;
    if(GF.to)chips+=`<span class="gf-chip">To: <strong>${GF.to}</strong><span class="gf-chip-x" onclick="removeGFVal('to','')">×</span></span>`;
    if(chips)chips+=`<span class="gf-chip gf-chip-clear" onclick="clearAllGF()"><i class="fas fa-times"></i> Clear All</span>`;

    document.getElementById('gf-chips').innerHTML = chips;
    const chChips = document.getElementById('ch-chips');
    if (chChips) chChips.innerHTML = chips;
    const rptChips = document.getElementById('rpt-chips');
    if (rptChips) rptChips.innerHTML = chips;
    const rptSummary = document.getElementById('rpt-filter-summary');
    if (rptSummary) rptSummary.textContent = chips ? 'Filters applied' : 'No filters';
}

function removeGFVal(key, val) {
    if (key === 'status' || key === 'from' || key === 'to') {
        GF[key] = '';
        MS_SUFFIXES.forEach(s => {
            const el = document.getElementById('gf-' + key + s);
            if (el) el.value = '';
        });
    } else if (Array.isArray(GF[key])) {
        GF[key] = GF[key].filter(v=>v!==val);
        const msId = 'gf-'+key;
        MS_SUFFIXES.forEach(s => {
            const opts = document.getElementById(msId + s + '-opts');
            if (opts) {
                opts.querySelectorAll('input[type=checkbox]').forEach(c => {
                    if (c.value === val) c.checked = false;
                });
            }
            updateMSTrigger(msId, s);
        });
    }
    applyGF();
    if(document.getElementById('fs-ov') && document.getElementById('fs-ov').classList.contains('show')) {
        syncFSSidebarFromGF();
    }
}

function applyGFToPool(pool) {
    const from=GF.from?new Date(GF.from):null,to=GF.to?new Date(GF.to):null;
    if(from)from.setHours(0,0,0,0);if(to)to.setHours(23,59,59,999);
    return pool.filter(e=>{
        if(GF.dep.length&&!GF.dep.includes(e.deposit))return false;
        if(GF.dept.length&&!GF.dept.includes(e.dept))return false;
        if(GF.sub.length&&!GF.sub.includes(e.subdept))return false;
        if(GF.skill.length&&!GF.skill.includes(e.skill))return false;
        if(GF.grade.length&&!GF.grade.includes(e.grade))return false;
        if(GF.desig.length&&!GF.desig.includes(e.desig))return false;
        if(GF.ry.length&&!GF.ry.includes(String(e.rd.getFullYear())))return false;
        if(GF.rfy.length&&!GF.rfy.includes(e.fy))return false;
        if(from&&e.rd<from)return false;
        if(to&&e.rd>to)return false;
        if(GF.status){
            const s=getStatus(e);
            if(GF.status==='active'&&s!=='active')return false;
            if(GF.status==='retired'&&s!=='retired')return false;
            if(GF.status==='soon'&&s!=='soon'&&s!=='year')return false;
            if(GF.status==='year'&&s!=='year')return false;
        }
        return true;
    });
}

// ===================== INSIGHT STRIP =====================
function updateInsightStrip() {
    const pool=applyGFToPool([...DATA,...PAST]);
    if(!pool.length){document.getElementById('emp-insight-strip').style.display='none';return;}
    document.getElementById('emp-insight-strip').style.display='flex';
    const active=pool.filter(e=>e.rd>=TODAY);
    const d1=addYears(1),d10=addYears(10);
    const r1=active.filter(e=>e.rd<=d1).length,r10=active.filter(e=>e.rd<=d10).length;
    const sm={};pool.forEach(e=>{const k=e.skill||'N/A';if(!sm[k])sm[k]=0;sm[k]++;});
    const topSkill=Object.entries(sm).sort((a,b)=>b[1]-a[1])[0];
    const insights=[];
    if(r1>0)insights.push(`<strong>${fmtN(r1)}</strong> employees retire in 1 year.`);
    if(r10>0&&pool.length)insights.push(`<strong>${(r10/pool.length*100).toFixed(1)}%</strong> workforce retires in 10 years.`);
    if(topSkill)insights.push(`Top skill: <strong>${esc(topSkill[0])}</strong> (${fmtN(topSkill[1])}).`);
    const avg=pool.length?pool.reduce((s,e)=>s+e.age,0)/pool.length:0;
    if(avg)insights.push(`Average age: <strong>${avg.toFixed(1)} yrs</strong>.`);
    document.getElementById('emp-insight-text').innerHTML=insights.join(' &nbsp;·&nbsp; ');
}

// ===================== KPI CARDS =====================
function updateKPIs() {
    const pool=applyGFToPool([...DATA,...PAST]);
    const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const d1=addYears(1),d5=addYears(5),d10=addYears(10);
    const tm1=new Date(now.getFullYear(),now.getMonth(),1),tm2=new Date(now.getFullYear(),now.getMonth()+1,0);
    const nm1=new Date(now.getFullYear(),now.getMonth()+1,1),nm2=new Date(now.getFullYear(),now.getMonth()+2,0);
    document.getElementById('k-this-month').textContent=pool.filter(e=>e.rd>=tm1&&e.rd<=tm2).length;
    document.getElementById('k-next-month').textContent=pool.filter(e=>e.rd>=nm1&&e.rd<=nm2).length;
    document.getElementById('k-3m').textContent=active.filter(e=>e.rd<=addMonths(3)).length;
    document.getElementById('k-1y').textContent=active.filter(e=>e.rd<=d1).length;
    document.getElementById('k-5y').textContent=active.filter(e=>e.rd<=d5).length;
    document.getElementById('k-10y').textContent=active.filter(e=>e.rd<=d10).length;
    document.getElementById('k-retired').textContent=pool.filter(e=>e.rd<now).length;
    document.getElementById('k-active').textContent=active.length;
    const aa=pool.length?pool.reduce((s,e)=>s+e.age,0)/pool.length:0;
    document.getElementById('k-avg-age').textContent=pool.length?aa.toFixed(1)+' yrs':'—';
    updateInsightStrip();
}

function openKPIDrill(type) {
    const pool=applyGFToPool([...DATA,...PAST]);
    const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const tm1=new Date(now.getFullYear(),now.getMonth(),1),tm2=new Date(now.getFullYear(),now.getMonth()+1,0);
    const nm1=new Date(now.getFullYear(),now.getMonth()+1,1),nm2=new Date(now.getFullYear(),now.getMonth()+2,0);
    let list=[];
    const labels={thisMonth:'Retiring This Month',nextMonth:'Retiring Next Month','3m':'Retiring Within 3 Months','1y':'Retiring Within 1 Year','5y':'Retiring Within 5 Years','10y':'Retiring Within 10 Years',retired:'Already Retired',active:'Active Employees',age:'Age Analysis'};
    if(type==='thisMonth')list=pool.filter(e=>e.rd>=tm1&&e.rd<=tm2);
    else if(type==='nextMonth')list=pool.filter(e=>e.rd>=nm1&&e.rd<=nm2);
    else if(type==='3m')list=active.filter(e=>e.rd<=addMonths(3));
    else if(type==='1y')list=active.filter(e=>e.rd<=addYears(1));
    else if(type==='5y')list=active.filter(e=>e.rd<=addYears(5));
    else if(type==='10y')list=active.filter(e=>e.rd<=addYears(10));
    else if(type==='retired')list=pool.filter(e=>e.rd<now);
    else if(type==='active')list=active;
    else if(type==='age')list=pool;
    openDrill(labels[type]||type,list);
}

// ===================== EMPLOYEE TABLE =====================
function getTableData() {
    const q=(document.getElementById('srch').value||'').toLowerCase();
    const dept=document.getElementById('lf-d').value;
    const gender=document.getElementById('lf-g').value;
    const status=document.getElementById('lf-status').value;
    let pool=applyGFToPool([...DATA,...PAST]);
    return pool.filter(e=>{
        if(q&&!(e.name+e.empid+e.dept+e.desig+e.grade+e.skill).toLowerCase().includes(q))return false;
        if(dept&&e.dept!==dept)return false;
        if(gender&&e.gender!==gender)return false;
        const s=getStatus(e);
        if(status==='active'&&s!=='active')return false;
        if(status==='retired'&&s!=='retired')return false;
        if(status==='soon'&&s!=='soon'&&s!=='year')return false;
        return true;
    });
}

const COLS_DEF=[
    {k:'cb',lbl:'cb',sk:null,r:(e,i)=>`<td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" class="row-cb" data-key="${esc(empKey(e))}" onchange="onRowCBChange(this)" ${SELECTED_IDS.has(empKey(e))?'checked':''}></td>`},
    {k:'#',lbl:'#',sk:null,r:(e,i)=>`<td class="cm">${i+1}</td>`},
    {k:'name',lbl:'Name',sk:'name',r:e=>`<td class="cn">${esc(e.name)}</td>`},
    {k:'empid',lbl:'Emp ID',sk:'empid',r:e=>`<td>${e.empid?`<span class="empid-badge">${esc(e.empid)}</span>`:'—'}</td>`},
    {k:'status',lbl:'Status',sk:'st',r:e=>`<td>${statusBadge(e)}</td>`},
    {k:'desig',lbl:'Designation',sk:'desig',r:e=>`<td>${esc(e.desig)}</td>`},
    {k:'grade',lbl:'Grade',sk:'grade',r:e=>`<td><span class="cdept">${esc(e.grade)}</span></td>`},
    {k:'dept',lbl:'Department',sk:'dept',r:e=>`<td><span class="cdept">${esc(e.dept)}</span></td>`},
    {k:'subdept',lbl:'Sub-Dept',sk:'subdept',r:e=>`<td style="color:var(--tx2)">${esc(e.subdept)}</td>`},
    {k:'skill',lbl:'Skill',sk:'skill',r:e=>`<td style="color:var(--sk2)">${esc(e.skill)}</td>`},
    {k:'dob',lbl:'DOB',sk:'dob',r:e=>`<td class="cm">${fmtDate(e.dob)}</td>`},
    {k:'rd',lbl:'Retirement Date',sk:'rd',r:e=>`<td class="rdate ${e.st||'past'}">${fmtDate(e.rd)}</td>`},
    {k:'fy',lbl:'FY',sk:'fyYear',r:e=>`<td><span class="fy-chip">${esc(e.fy)}</span></td>`}
];

function onRowCBChange(cb){const key=cb.dataset.key;if(cb.checked)SELECTED_IDS.add(key);else SELECTED_IDS.delete(key);updateSelectionToolbar();updateHeaderCB();}
function onHeaderCBChange(cb){document.querySelectorAll('#tbody .row-cb').forEach(c=>{c.checked=cb.checked;const key=c.dataset.key;if(cb.checked)SELECTED_IDS.add(key);else SELECTED_IDS.delete(key);});updateSelectionToolbar();}
function updateHeaderCB(){const allCBs=document.querySelectorAll('#tbody .row-cb');const hCB=document.getElementById('th-master-cb');if(!hCB||!allCBs.length)return;const checkedCount=[...allCBs].filter(c=>c.checked).length;hCB.indeterminate=checkedCount>0&&checkedCount<allCBs.length;hCB.checked=checkedCount===allCBs.length&&allCBs.length>0;}
function updateSelectionToolbar(){const tb=document.getElementById('sel-toolbar'),cnt=document.getElementById('sel-count');if(!tb)return;if(SELECTED_IDS.size>0){tb.style.display='flex';cnt.textContent=SELECTED_IDS.size;}else{tb.style.display='none';}}
function selectAllVisible(){document.querySelectorAll('#tbody .row-cb').forEach(c=>{c.checked=true;SELECTED_IDS.add(c.dataset.key);});updateSelectionToolbar();updateHeaderCB();showToast(`${document.querySelectorAll('#tbody .row-cb').length} visible rows selected`,'fa-check-square');}
function selectAllPages(){FILTERED_CACHE.forEach(e=>SELECTED_IDS.add(empKey(e)));buildPage();updateSelectionToolbar();showToast(`${FILTERED_CACHE.length} records selected`,'fa-check-square');}
function clearSelection(){SELECTED_IDS.clear();document.querySelectorAll('#tbody .row-cb').forEach(c=>c.checked=false);updateSelectionToolbar();updateHeaderCB();showToast('Selection cleared','fa-times');}
function exportSelected(type){const all=[...DATA,...PAST];const selected=all.filter(e=>SELECTED_IDS.has(empKey(e)));if(!selected.length){alert('No employees selected.');return;}_exportEmpData(selected,type,'NMDC_Selected');}
function viewSelectedDrill(){const all=[...DATA,...PAST];const selected=all.filter(e=>SELECTED_IDS.has(empKey(e)));if(!selected.length){alert('No employees selected.');return;}openDrill(`Selected Employees (${selected.length})`,selected);}

function renderTable(){CUR_PAGE=1;FILTERED_CACHE=applySort(getTableData());document.getElementById('rcount').textContent=`${FILTERED_CACHE.length} records`;buildPage();}
function applySort(arr){if(!SORT_COL)return arr;return[...arr].sort((a,b)=>{let av=a[SORT_COL],bv=b[SORT_COL];if(av instanceof Date&&bv instanceof Date)return(av-bv)*SORT_DIR;if(typeof av==='number'&&typeof bv==='number')return(av-bv)*SORT_DIR;return String(av||'').localeCompare(String(bv||''))*SORT_DIR;});}
function buildPage(){const tot=FILTERED_CACHE.length,sz=PAGE_SIZE===0?tot:PAGE_SIZE,pgs=sz?Math.ceil(tot/sz):1;CUR_PAGE=Math.min(CUR_PAGE,pgs||1);const st=(CUR_PAGE-1)*sz,slice=PAGE_SIZE===0?FILTERED_CACHE:FILTERED_CACHE.slice(st,st+sz);document.getElementById('thead').innerHTML=COLS_DEF.map((c,ci)=>{if(c.k==='cb')return`<th class="cb-col"><input type="checkbox" id="th-master-cb" onchange="onHeaderCBChange(this)" title="Select all visible"></th>`;if(!c.sk)return`<th>${c.lbl}</th>`;const cls=SORT_COL===c.sk?(SORT_DIR===1?'sortable asc':'sortable desc'):'sortable';return`<th class="${cls}" onclick="sortBy('${c.sk}')">${c.lbl}<span class="si"></span></th>`;}).join('');document.getElementById('tbody').innerHTML=slice.length?slice.map((e,i)=>{const cells=COLS_DEF.map((c,ci)=>ci===0?c.r(e,st+i):ci===1?c.r(e,st+i):c.r(e)).join('');const ej=JSON.stringify(e).replace(/'/g,"&#39;");const isSelected=SELECTED_IDS.has(empKey(e));return`<tr class="${isSelected?'row-selected':''}" onclick='openEmpPopup(${ej})'>${cells}</tr>`;}).join(''):`<tr><td colspan="${COLS_DEF.length}" style="text-align:center;padding:28px;color:var(--tx3)">No records</td></tr>`;updateHeaderCB();buildPag(tot,sz,pgs);}
function sortBy(col){if(SORT_COL===col)SORT_DIR*=-1;else{SORT_COL=col;SORT_DIR=1;}FILTERED_CACHE=applySort(FILTERED_CACHE);buildPage();}
function buildPag(tot,sz,pgs){const bar=document.getElementById('pag-bar');if(!tot||(pgs<=1&&PAGE_SIZE!==0)){bar.style.display='none';return;}bar.style.display='flex';const s=PAGE_SIZE===0?1:(CUR_PAGE-1)*sz+1,en=PAGE_SIZE===0?tot:Math.min(CUR_PAGE*sz,tot);document.getElementById('pag-info').textContent=`${s}–${en} of ${tot}`;let btns=`<button class="pag-btn" onclick="goPage(${CUR_PAGE-1})" ${CUR_PAGE===1?'disabled':''}>‹</button>`;const range=(c,t)=>{if(t<=7)return Array.from({length:t},(_,i)=>i+1);let r=[1];if(c>3)r.push('…');for(let p=Math.max(2,c-1);p<=Math.min(t-1,c+1);p++)r.push(p);if(c<t-2)r.push('…');r.push(t);return r;};range(CUR_PAGE,pgs).forEach(p=>{if(p==='…')btns+=`<span style="padding:0 3px;color:var(--tx3)">…</span>`;else btns+=`<button class="pag-btn ${p===CUR_PAGE?'active':''}" onclick="goPage(${p})">${p}</button>`;});btns+=`<button class="pag-btn" onclick="goPage(${CUR_PAGE+1})" ${CUR_PAGE===pgs?'disabled':''}>›</button>`;document.getElementById('pag-btns').innerHTML=btns;}
function goPage(p){const pgs=Math.ceil(FILTERED_CACHE.length/(PAGE_SIZE||FILTERED_CACHE.length||1));if(p<1||p>pgs)return;CUR_PAGE=p;buildPage();}
function onPageSzChg(sel){PAGE_SIZE=parseInt(sel.value);CUR_PAGE=1;buildPage();}
function lReset(){['lf-d','lf-g','lf-status'].forEach(id=>document.getElementById(id).value='');document.getElementById('srch').value='';SORT_COL=null;SORT_DIR=1;CUR_PAGE=1;renderTable();}

// ===================== DRILL MODAL =====================
let _drillDCH = {};
function openDrill(title, list) {
    DRILL_ALL=list;DRILL_DATA=[...list];DRILL_TITLE=title;
    DRILL_SELECTED_IDS.clear();
    document.getElementById('drill-title').textContent=title;
    buildDrillKPIs(list);buildDrillCharts(list);renderDrillTable(DRILL_DATA);
    document.getElementById('drill-ov').style.display='flex';
    updateDrillSelectionToolbar();
}
function buildDrillKPIs(list) {
    const now=TODAY,active=list.filter(e=>e.rd>=now);
    const d1=addYears(1),d5=addYears(5),d10=addYears(10);
    const avg=list.length?list.reduce((s,e)=>s+e.age,0)/list.length:0;
    const kpis=[
        {n:fmtN(list.length),l:'Total',c:'var(--ind)'},{n:fmtN(active.filter(e=>e.rd<=addMonths(3)).length),l:'Retiring 3M',c:'#dc2626'},
        {n:fmtN(active.filter(e=>e.rd<=d1).length),l:'Retiring 1Y',c:'#d97706'},{n:fmtN(active.filter(e=>e.rd<=d5).length),l:'Retiring 5Y',c:'#ca8a04'},
        {n:fmtN(active.filter(e=>e.rd<=d10).length),l:'Retiring 10Y',c:'#059669'},{n:fmtN(list.filter(e=>e.rd<now).length),l:'Retired',c:'var(--sk)'},
        {n:fmtN(active.length),l:'Active',c:'var(--em)'},{n:avg.toFixed(1),l:'Avg Age',c:'#7c3aed'}
    ];
    document.getElementById('drill-kpis').innerHTML=kpis.map(k=>`<div class="drill-kpi"><div class="drill-kpi-num" style="color:${k.c}">${k.n}</div><div class="drill-kpi-lbl">${k.l}</div></div>`).join('');
}
function buildDrillCharts(list) {
    Object.values(_drillDCH).forEach(c=>{try{c.destroy();}catch(e){}});_drillDCH={};
    document.getElementById('drill-chart-panel').style.display='none';
}
function onDrillCBChange(cb){const key=cb.dataset.key;if(cb.checked)DRILL_SELECTED_IDS.add(key);else DRILL_SELECTED_IDS.delete(key);updateDrillSelectionToolbar();updateDrillHeaderCB();}
function onDrillHeaderCBChange(cb){document.querySelectorAll('#drill-tbody .row-cb').forEach(c=>{c.checked=cb.checked;const key=c.dataset.key;if(cb.checked)DRILL_SELECTED_IDS.add(key);else DRILL_SELECTED_IDS.delete(key);});updateDrillSelectionToolbar();}
function updateDrillHeaderCB(){const allCBs=document.querySelectorAll('#drill-tbody .row-cb');const hCB=document.getElementById('drill-master-cb');if(!hCB||!allCBs.length)return;const checkedCount=[...allCBs].filter(c=>c.checked).length;hCB.indeterminate=checkedCount>0&&checkedCount<allCBs.length;hCB.checked=checkedCount===allCBs.length&&allCBs.length>0;}
function updateDrillSelectionToolbar(){const tb=document.getElementById('drill-sel-toolbar'),cnt=document.getElementById('drill-sel-count');if(!tb)return;if(DRILL_SELECTED_IDS.size>0){tb.style.display='flex';cnt.textContent=DRILL_SELECTED_IDS.size;}else{tb.style.display='none';}}
function selectAllDrillVisible(){document.querySelectorAll('#drill-tbody .row-cb').forEach(c=>{c.checked=true;DRILL_SELECTED_IDS.add(c.dataset.key);});updateDrillSelectionToolbar();updateDrillHeaderCB();showToast(`${DRILL_SELECTED_IDS.size} selected`,'fa-check-square');}
function clearDrillSelection(){DRILL_SELECTED_IDS.clear();document.querySelectorAll('#drill-tbody .row-cb').forEach(c=>c.checked=false);updateDrillSelectionToolbar();updateDrillHeaderCB();showToast('Selection cleared','fa-times');}
function exportDrillSelected(type){const selected=DRILL_ALL.filter(e=>DRILL_SELECTED_IDS.has(empKey(e)));if(!selected.length){alert('No employees selected.');return;}_exportEmpData(selected,type,'NMDC_DrillSelected');}

function renderDrillTable(list) {
    document.getElementById('drill-rc').textContent=`${list.length} records`;
    document.getElementById('drill-thead').innerHTML=`<tr><th class="cb-col"><input type="checkbox" id="drill-master-cb" onchange="onDrillHeaderCBChange(this)" title="Select all"></th><th>#</th><th>Name</th><th>Emp ID</th><th>Status</th><th>Grade</th><th>Designation</th><th>Skill</th><th>Department</th><th>Sub-Dept</th><th>DOB</th><th>Retirement Date</th><th>FY</th></tr>`;
    document.getElementById('drill-tbody').innerHTML=list.map((e,i)=>{
        const ej=JSON.stringify(e).replace(/'/g,"&#39;");const key=empKey(e);const isChecked=DRILL_SELECTED_IDS.has(key);
        return`<tr class="${isChecked?'row-selected':''}">
            <td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" class="row-cb" data-key="${esc(key)}" onchange="onDrillCBChange(this)" ${isChecked?'checked':''}></td>
            <td class="cm">${i+1}</td>
            <td class="cn" onclick='openEmpPopup(${ej})' style="cursor:pointer">${esc(e.name)}</td>
            <td onclick='openEmpPopup(${ej})' style="cursor:pointer">${e.empid?`<span class="empid-badge">${esc(e.empid)}</span>`:'—'}</td>
            <td onclick='openEmpPopup(${ej})' style="cursor:pointer">${statusBadge(e)}</td>
            <td onclick='openEmpPopup(${ej})' style="cursor:pointer"><span class="cdept">${esc(e.grade)}</span></td>
            <td onclick='openEmpPopup(${ej})' style="cursor:pointer">${esc(e.desig)}</td>
            <td onclick='openEmpPopup(${ej})' style="cursor:pointer">${esc(e.skill)}</td>
            <td onclick='openEmpPopup(${ej})' style="cursor:pointer"><span class="cdept">${esc(e.dept)}</span></td>
            <td style="color:var(--tx2)" onclick='openEmpPopup(${ej})'>${esc(e.subdept)}</td>
            <td class="cm" onclick='openEmpPopup(${ej})'>${fmtDate(e.dob)}</td>
            <td class="rdate ${e.st||'past'}" onclick='openEmpPopup(${ej})'>${fmtDate(e.rd)}</td>
            <td onclick='openEmpPopup(${ej})'><span class="fy-chip">${esc(e.fy)}</span></td>
        </tr>`;
    }).join('');
    updateDrillHeaderCB();
}

function exportDrill(type) {
    if(!DRILL_DATA.length){alert('No data');return;}
    if(type==='xlsx'){
        const pool=applyGFToPool([...DATA,...PAST]);
        const now=TODAY,active=pool.filter(e=>e.rd>=now);
        const d1=addYears(1),d5=addYears(5),d10=addYears(10);
        const tm1=new Date(now.getFullYear(),now.getMonth(),1),tm2=new Date(now.getFullYear(),now.getMonth()+1,0);
        const nm1=new Date(now.getFullYear(),now.getMonth()+1,1),nm2=new Date(now.getFullYear(),now.getMonth()+2,0);
        const summary=[['Category','Count'],['Retiring This Month',pool.filter(e=>e.rd>=tm1&&e.rd<=tm2).length],['Retiring Next Month',pool.filter(e=>e.rd>=nm1&&e.rd<=nm2).length],['Within 3 Months',active.filter(e=>e.rd<=addMonths(3)).length],['Within 1 Year',active.filter(e=>e.rd<=d1).length],['Within 5 Years',active.filter(e=>e.rd<=d5).length],['Within 10 Years',active.filter(e=>e.rd<=d10).length],[],['Current View — '+DRILL_TITLE,DRILL_DATA.length]];
        const detailHdrs=['Name','Emp ID','Grade','Designation','Skill','Department','Sub-Dept','Deposit','DOB','Retirement Date','FY','Gender','Status'];
        const detailRows=DRILL_DATA.map(e=>[e.name,e.empid,e.grade,e.desig,e.skill,e.dept,e.subdept,e.deposit,fmtDate(e.dob),fmtDate(e.rd),e.fy,e.gender,getStatus(e)]);
        const aoa=[...summary,[],detailHdrs,...detailRows];
        const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=[{wch:32},{wch:12},{wch:10},{wch:22},{wch:18},{wch:22},{wch:18},{wch:12},{wch:14},{wch:16},{wch:12},{wch:8},{wch:12}];
        const wb=XLSX.utils.book_new();const sheetName=DRILL_TITLE.slice(0,31).replace(/[:\\\/\?\*\[\]]/g,'');
        XLSX.utils.book_append_sheet(wb,ws,sheetName||'Detail');XLSX.writeFile(wb,'NMDC_'+(sheetName.replace(/\s+/g,'_')||'Detail')+'.xlsx');showToast('Excel exported!','fa-file-excel');
    }else{window.print();showToast('Sent to printer','fa-print');}
}
function closeDrill(){document.getElementById('drill-ov').style.display='none';DRILL_SELECTED_IDS.clear();}
function closeDrillOv(e){if(e.target.id==='drill-ov')closeDrill();}

// ===================== EMPLOYEE POPUP =====================
function openEmpPopup(emp) {
    if(typeof emp==='string')emp=JSON.parse(emp);
    if(emp.dob&&!(emp.dob instanceof Date))emp.dob=new Date(emp.dob);
    if(emp.rd&&!(emp.rd instanceof Date))emp.rd=new Date(emp.rd);
    if(emp.b60&&!(emp.b60 instanceof Date))emp.b60=new Date(emp.b60);
    if(!emp.b60&&emp.dob)emp.b60=get60th(emp.dob);
    const s=getStatus(emp),stL={retired:'Retired',soon:'Retiring Soon',year:'Retiring This Year',active:'Active'}[s]||'Active';
    const stC={retired:'var(--sk)',soon:'var(--ro)',year:'var(--am)',active:'var(--em)'}[s];
    const init=emp.name.split(' ').map(w=>w[0]||'').join('').slice(0,2).toUpperCase()||'?';
    const cols=['#1E3A8A','#065F46','#4C1D95','#92400E','#0C4A6E','#991B1B'];
    const col=cols[(emp.name.charCodeAt(0)||0)%cols.length];
    document.getElementById('ep-avatar').style.cssText=`background:${col}22;color:${col}`;
    document.getElementById('ep-avatar').textContent=init;
    document.getElementById('ep-name').textContent=emp.name;
    document.getElementById('ep-sub').textContent=`${emp.desig} · ${emp.dept}`;
    document.getElementById('ep-calc').innerHTML=retCalcExpl(emp.dob,emp.b60,emp.rd);
    document.getElementById('ep-grid').innerHTML=[
        ['Date of Birth',fmtDate(emp.dob)],['Age',emp.age+' years'],['60th Birthday',fmtDate(emp.b60)],
        ['Retirement Date',`<span style="color:var(--ind)">${fmtDate(emp.rd)}</span>`],
        ['Financial Year',`<span class="fy-chip">${esc(emp.fy)}</span>`],['Grade',esc(emp.grade)],
        ['Skill',esc(emp.skill)],['Deposit',esc(emp.deposit)],
        ['Department',esc(emp.dept)],['Sub-Department',esc(emp.subdept)],
        ['Gender',emp.gender==='M'?'Male':emp.gender==='F'?'Female':'Other'],
        ['Status',`<span style="color:${stC}">${stL}</span>`]
    ].map(([l,v])=>`<div class="ep-field"><div class="ep-flbl">${l}</div><div class="ep-fval">${v}</div></div>`).join('');
    document.getElementById('emp-popup-ov').style.display='flex';
}
function closeEmpOv(e){if(e.target.id==='emp-popup-ov')document.getElementById('emp-popup-ov').style.display='none';}

// ===================== CHARTS =====================
const PAL=['#1E3A8A','#065F46','#4C1D95','#92400E','#991B1B','#0C4A6E','#2563EB','#059669','#7C3AED','#D97706','#DC2626','#0284C7','#be185d','#0f766e','#7e22ce','#854d0e','#0369a1','#047857'];
const CBASE={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{font:{family:"'Source Sans 3',sans-serif",size:11},padding:8,boxWidth:11}},tooltip:{backgroundColor:'#1A1D2E',titleColor:'#fff',bodyColor:'#BFC5D6',borderColor:'#2563EB',borderWidth:1,padding:10,cornerRadius:8}}};

function dCh(id){if(CH[id]){CH[id].destroy();delete CH[id];}}

function drawAllCharts() {
    if(!DATA.length&&!PAST.length)return;
    const pool=applyGFToPool([...DATA,...PAST]);
    drawAgeChart(pool);drawDeptChart(pool);drawMonthChart(pool);
    drawDesigChart(pool);drawSkillChart(pool);drawFYChart(pool);drawGradeChart(pool);
}

function drawAgeChart(data) {
    const bands=[{lo:20,hi:30,lbl:'20-30',c:'#6366f1'},{lo:31,hi:40,lbl:'31-40',c:'#8b5cf6'},{lo:41,hi:50,lbl:'41-50',c:'#3b82f6'},{lo:51,hi:55,lbl:'51-55',c:'#f59e0b'},{lo:56,hi:59,lbl:'56-59',c:'#ef4444'},{lo:60,hi:999,lbl:'60+',c:'#991b1b'}];
    const cnts=bands.map(b=>({...b,cnt:data.filter(e=>e.age>=b.lo&&e.age<=b.hi).length}));
    const tot=data.length||1;
    document.getElementById('age-band-grid').innerHTML=cnts.map(b=>`<div class="age-band-card" style="cursor:default"><div class="age-band-num" style="color:${b.c}">${b.cnt}</div><div class="age-band-lbl">${b.lbl} yrs</div><div class="age-band-bar" style="background:${b.c}22"><div style="width:${Math.min(b.cnt/tot*100,100).toFixed(0)}%;height:100%;background:${b.c};border-radius:2px"></div></div></div>`).join('');
    dCh('ca-age-dist');
    CH['ca-age-dist']=new Chart(document.getElementById('ca-age-dist'),{type:'bar',data:{labels:cnts.map(b=>b.lbl),datasets:[{label:'Employees',data:cnts.map(b=>b.cnt),backgroundColor:cnts.map(b=>b.c+'bb'),borderColor:cnts.map(b=>b.c),borderWidth:1.5,borderRadius:5}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}}}});
}
function drawDeptChart(data){dCh('ca-dept');const m={};data.forEach(e=>m[e.dept]=(m[e.dept]||0)+1);const s=Object.entries(m).sort((a,b)=>b[1]-a[1]);CH['ca-dept']=new Chart(document.getElementById('ca-dept'),{type:'bar',data:{labels:s.map(d=>d[0]),datasets:[{label:'Count',data:s.map(d=>d[1]),backgroundColor:s.map((_,i)=>PAL[i%PAL.length]+'bb'),borderColor:s.map((_,i)=>PAL[i%PAL.length]),borderWidth:1.5,borderRadius:4}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true},x:{ticks:{maxRotation:30}}}}});}
function drawMonthChart(data){dCh('ca-month');const mm=new Map();data.forEach(e=>{if(!mm.has(e.rdMS))mm.set(e.rdMS,{lbl:e.rdMonth,cnt:0,sk:e.rd.getFullYear()*12+e.rd.getMonth()});mm.get(e.rdMS).cnt++;});const s=[...mm.entries()].sort((a,b)=>a[1].sk-b[1].sk);const now=TODAY;const barColors=s.map(x=>{const yr=parseInt(x[0].split('-')[0]),mo=parseInt(x[0].split('-')[1])-1;const d=new Date(yr,mo,1);if(d<new Date(now.getFullYear(),now.getMonth(),1))return'#94a3b8bb';if(yr===now.getFullYear())return'#d97706bb';if(yr===now.getFullYear()+1)return'#2563ebbb';return'#065f46bb';});CH['ca-month']=new Chart(document.getElementById('ca-month'),{type:'bar',data:{labels:s.map(x=>x[1].lbl),datasets:[{label:'Retirees',data:s.map(x=>x[1].cnt),backgroundColor:barColors,borderColor:barColors.map(c=>c.replace(/bb$/,'ff')),borderWidth:1,borderRadius:3}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true},x:{ticks:{maxRotation:45,font:{size:10}}}}}});}
function drawDesigChart(data){dCh('ca-desig');const m={};data.forEach(e=>m[e.desig]=(m[e.desig]||0)+1);const s=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,8);CH['ca-desig']=new Chart(document.getElementById('ca-desig'),{type:'doughnut',data:{labels:s.map(d=>d[0]),datasets:[{data:s.map(d=>d[1]),backgroundColor:PAL.slice(0,s.length),borderWidth:2,borderColor:DARK?'#1a1d2e':'#fff',hoverOffset:8}]},options:{...CBASE,cutout:'55%',plugins:{...CBASE.plugins,legend:{position:'right'}}}});}
function drawSkillChart(data){dCh('ca-skill');const m={};data.forEach(e=>{const k=e.skill!=='N/A'?e.skill:e.desig;m[k]=(m[k]||0)+1;});const s=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,10);CH['ca-skill']=new Chart(document.getElementById('ca-skill'),{type:'bar',data:{labels:s.map(x=>x[0]),datasets:[{label:'Count',data:s.map(x=>x[1]),backgroundColor:'#7C3AEDbb',borderColor:'#7C3AED',borderRadius:4}]},options:{...CBASE,indexAxis:'y',plugins:{...CBASE.plugins,legend:{display:false}},scales:{x:{beginAtZero:true}}}});}
function drawFYChart(data){dCh('ca-fy');const m={};data.forEach(e=>m[e.fy]=(m[e.fy]||0)+1);const s=Object.entries(m).sort((a,b)=>getFYYear(a[0])-getFYYear(b[0]));CH['ca-fy']=new Chart(document.getElementById('ca-fy'),{type:'bar',data:{labels:s.map(y=>y[0]),datasets:[{label:'Retirees',data:s.map(y=>y[1]),backgroundColor:s.map((_,i)=>PAL[i%PAL.length]+'bb'),borderRadius:4}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true}}}});}
function drawGradeChart(data){dCh('ca-grade');const m={};data.forEach(e=>m[e.grade]=(m[e.grade]||0)+1);const s=Object.entries(m).sort((a,b)=>{const na=parseFloat((a[0].match(/\d+(\.\d+)?/)||[0])[0]);const nb=parseFloat((b[0].match(/\d+(\.\d+)?/)||[0])[0]);if(!isNaN(na)&&!isNaN(nb))return nb-na;return b[0].localeCompare(a[0]);}).slice(0,12);CH['ca-grade']=new Chart(document.getElementById('ca-grade'),{type:'bar',data:{labels:s.map(g=>g[0]),datasets:[{label:'Count',data:s.map(g=>g[1]),backgroundColor:'#d9770666',borderColor:'#d97706',borderRadius:4}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true},x:{ticks:{maxRotation:30}}}}});}

function showDrillBy(type,label){
    const all=applyGFToPool([...DATA,...PAST]);let list=[];
    if(type==='dept')list=all.filter(e=>e.dept===label);
    else if(type==='desig')list=all.filter(e=>e.desig===label);
    else if(type==='grade')list=all.filter(e=>e.grade===label);
    else if(type==='skill')list=all.filter(e=>e.skill===label);
    else if(type==='month')list=all.filter(e=>e.rdMonth===label);
    else if(type==='fy')list=all.filter(e=>e.fy===label);
    else if(type==='ageBand'){const[lo,hi]=label.split('-').map(Number);const isPlus=label.endsWith('+');list=all.filter(e=>isPlus?e.age>=60:e.age>=lo&&e.age<=hi);}
    openDrill(`${type}: ${label} (${list.length})`,list);
}

// ===================== TAB NAVIGATION =====================
function GT(btn) {
    document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
    document.querySelectorAll('.pg').forEach(p=>p.classList.remove('on'));
    btn.classList.add('on');
    document.getElementById('pg-'+btn.dataset.p).classList.add('on');
    if(btn.dataset.p==='charts'&&(DATA.length||PAST.length))setTimeout(drawAllCharts,60);
}

// ===================== FULLSCREEN CHART MODAL =====================
function getFSBasePool() {
    const gfPool = applyGFToPool([...DATA,...PAST]);
    if (!FS_LOCAL.timePeriod) return gfPool;
    const now=TODAY,curYear=now.getFullYear();
    if(FS_LOCAL.timePeriod==='thisYear'){const start=new Date(curYear,0,1);const end=new Date(curYear,11,31);return gfPool.filter(e=>e.rd>=start&&e.rd<=end);}
    else if(FS_LOCAL.timePeriod==='nextYear'){const start=new Date(curYear+1,0,1);const end=new Date(curYear+1,11,31);return gfPool.filter(e=>e.rd>=start&&e.rd<=end);}
    else if(FS_LOCAL.timePeriod==='next5y'){const start=new Date(now);const end=addYears(5);return gfPool.filter(e=>e.rd>=start&&e.rd<=end);}
    return gfPool;
}

function getFSDisplayPool() {
    const base=getFSBasePool();
    if(!FS_LOCAL.clickedType||!FS_LOCAL.clickedValues.length)return base;
    const {clickedType:type,clickedValues:values}=FS_LOCAL;
    return base.filter(e=>{
        if(type==='dept')return values.includes(e.dept);
        if(type==='desig')return values.includes(e.desig);
        if(type==='grade')return values.includes(e.grade);
        if(type==='skill')return values.includes(e.skill);
        if(type==='month')return values.includes(e.rdMonth);
        if(type==='fy')return values.includes(e.fy);
        if(type==='deposit')return values.includes(e.deposit);
        if(type==='ageBand'){return values.some(lbl=>{if(lbl==='60+')return e.age>=60;const[lo,hi]=lbl.split('-').map(Number);return e.age>=lo&&e.age<=hi;});}
        return true;
    });
}

function fsCategoryClick(type,value) {
    if(FS_LOCAL.clickedType===type){
        const idx=FS_LOCAL.clickedValues.indexOf(value);
        if(idx>=0){FS_LOCAL.clickedValues=[];FS_LOCAL.clickedType=null;}
        else{FS_LOCAL.clickedValues=[value];}
    }else{FS_LOCAL.clickedType=type;FS_LOCAL.clickedValues=[value];}
    refreshFSDetailOnly();
}

function clearFSLocalClick(){FS_LOCAL.clickedType=null;FS_LOCAL.clickedValues=[];refreshFSDetailOnly();}

function refreshFSDetailOnly(){const pool=getFSDisplayPool();updateFSStats(pool);updateFSTable(pool);updateFSInsights(pool);renderFSClickBanner();}

function setFSTimePeriod(period) {
    FS_LOCAL.timePeriod = period || '';
    FS_LOCAL.clickedType=null;FS_LOCAL.clickedValues=[];
    updateFSTimeBtns();
    renderFSClickBanner();
    refreshFS();
}

function updateFSTimeBtns() {
    ['thisYear','nextYear','next5y'].forEach(p=>{
        const btn=document.getElementById(`fs-time-${p}`);
        if(btn)btn.classList.toggle('fs-time-btn-active',FS_LOCAL.timePeriod===p);
    });
    const clearBtn=document.getElementById('fs-time-clear');
    if(clearBtn)clearBtn.style.display=FS_LOCAL.timePeriod?'inline-flex':'none';
}

function openFS(type, title) {
    FS_TYPE=type;
    FS_LOCAL.timePeriod='';FS_LOCAL.clickedType=null;FS_LOCAL.clickedValues=[];
    document.getElementById('fs-title').textContent=title;
    document.getElementById('fs-ov').classList.add('show');

    // Sync FS sidebar inputs to match the current global GF state
    syncFSSidebarFromGF();

    // Highlight the active visual nav item in left sidebar
    document.querySelectorAll('.fs-nav-item').forEach(item => {
        item.classList.toggle('active', item.dataset.type === type);
    });

    updateFSTimeBtns();
    updateFSAppliedFilters();
    renderFSClickBanner();
    refreshFS();
}

// Sync the FS popup sidebar filter controls to reflect the current GF state
function syncFSSidebarFromGF() {
    // Sync status, from, to date inputs
    const stEl = document.getElementById('gf-status-fs');
    const frEl = document.getElementById('gf-from-fs');
    const toEl = document.getElementById('gf-to-fs');
    if (stEl) stEl.value = GF.status || '';
    if (frEl) frEl.value = GF.from || '';
    if (toEl) toEl.value = GF.to || '';

    // Sync MS checkboxes and trigger labels for -fs suffix
    ['gf-dep','gf-dept','gf-sub','gf-skill','gf-grade','gf-desig','gf-ry','gf-rfy'].forEach(id => {
        const data = MS_DATA[id];
        if (!data) return;
        const opts = document.getElementById(id + '-fs-opts');
        if (opts) {
            opts.querySelectorAll('input[type=checkbox]').forEach(c => {
                c.checked = data.selected.includes(c.value);
            });
        }
        updateMSTrigger(id, '-fs');
    });
}
function closeFS(){document.getElementById('fs-ov').classList.remove('show');if(FS_CH){FS_CH.destroy();FS_CH=null;}}
function closeFSOv(e){if(e.target.id==='fs-ov')closeFS();}

function getAppliedFiltersText() {
    const parts=[];
    if(GF.dep.length)parts.push({label:'Deposit',values:GF.dep});
    if(GF.dept.length)parts.push({label:'Department',values:GF.dept});
    if(GF.sub.length)parts.push({label:'Sub-Dept',values:GF.sub});
    if(GF.skill.length)parts.push({label:'Skill',values:GF.skill});
    if(GF.grade.length)parts.push({label:'Grade',values:GF.grade});
    if(GF.desig.length)parts.push({label:'Designation',values:GF.desig});
    if(GF.status)parts.push({label:'Status',values:[GF.status]});
    if(GF.ry.length)parts.push({label:'Ret. Year',values:GF.ry});
    if(GF.rfy.length)parts.push({label:'Ret. FY',values:GF.rfy});
    if(GF.from)parts.push({label:'From',values:[GF.from]});
    if(GF.to)parts.push({label:'To',values:[GF.to]});
    return parts;
}

function updateFSAppliedFilters() {
    const el=document.getElementById('fs-applied-filters');
    if(!el)return;
    const parts=getAppliedFiltersText();
    const timeLbl={thisYear:'This Year',nextYear:'Next Year',next5y:'Next 5 Years'}[FS_LOCAL.timePeriod];
    if(!parts.length&&!timeLbl){el.innerHTML='<span class="fs-no-filter">No filters applied — showing all data</span>';return;}
    let html=parts.map(p=>{const vals=p.values.map(v=>`<span class="fs-filter-chip">${esc(v)}</span>`).join(' ');return`<span class="fs-filter-group"><span class="fs-filter-key">${p.label}:</span>${vals}</span>`;}).join('');
    if(timeLbl)html+=`<span class="fs-filter-group"><span class="fs-filter-key">Period:</span><span class="fs-filter-chip fs-filter-chip-time">${timeLbl}</span></span>`;
    el.innerHTML=html;
}

function renderFSClickBanner() {
    const el=document.getElementById('fs-click-banner');
    if(!el)return;
    if(FS_LOCAL.clickedType&&FS_LOCAL.clickedValues.length){
        const typeLabel={dept:'Department',desig:'Designation',grade:'Grade',skill:'Skill',month:'Month',fy:'Financial Year',deposit:'Deposit',ageBand:'Age Band'}[FS_LOCAL.clickedType]||FS_LOCAL.clickedType;
        el.innerHTML=`<i class="fas fa-filter" style="color:var(--ind2)"></i> <strong>${typeLabel}:</strong> ${FS_LOCAL.clickedValues.map(v=>`<span style="background:var(--indb);border-radius:4px;padding:1px 7px;margin:0 2px;font-weight:700;color:var(--ind)">${esc(v)}</span>`).join('')} — detail panel filtered &nbsp;<button onclick="clearFSLocalClick()" style="background:var(--rol);border:1px solid var(--ro2)22;color:var(--ro2);border-radius:6px;padding:2px 9px;font-size:11px;font-weight:700;cursor:pointer">× Clear</button>`;
        el.style.display='flex';
    }else{el.style.display='none';}
}

function updateFSStats(pool) {
    const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const d1=addYears(1),d5=addYears(5),d10=addYears(10);
    const avg=pool.length?pool.reduce((s,e)=>s+e.age,0)/pool.length:0;
    document.getElementById('fs-stats').innerHTML=[
        {n:fmtN(pool.length),l:'Total',c:'var(--ind)'},{n:fmtN(active.filter(e=>e.rd<=d1).length),l:'Retiring 1Y',c:'#dc2626'},
        {n:fmtN(active.filter(e=>e.rd<=d5).length),l:'Retiring 5Y',c:'#d97706'},{n:fmtN(active.filter(e=>e.rd<=d10).length),l:'Retiring 10Y',c:'#059669'},
        {n:fmtN(pool.filter(e=>e.rd<now).length),l:'Retired',c:'var(--sk)'},{n:fmtN(active.length),l:'Active',c:'var(--em)'},
        {n:avg.toFixed(1),l:'Avg Age',c:'#7c3aed'},{n:fmtN(active.filter(e=>e.rd<=addMonths(3)).length),l:'Retiring 3M',c:'#ea580c'}
    ].map(s=>`<div class="fs-stat"><div class="fs-stat-num" style="color:${s.c}">${s.n}</div><div class="fs-stat-lbl">${s.l}</div></div>`).join('');
}

function updateFSTable(pool) {
    if(!FS_TYPE)return;
    const tableData=_buildFSTableDataFromPool(FS_TYPE,pool);
    FS_TABLE_DATA=tableData;
    document.getElementById('fs-thead').innerHTML=`<tr><th>${FS_TABLE_KEY}</th><th style="text-align:right">Count</th></tr>`;
    document.getElementById('fs-tbody').innerHTML=tableData.map(([k,v])=>`<tr><td class="cn">${esc(String(k))}</td><td style="text-align:right;font-family:var(--font-mono);font-weight:700">${fmtN(v)}</td></tr>`).join('');
}

function updateFSInsights(pool) {
    const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const items=[];
    if(FS_TABLE_DATA.length>0){const top=FS_TABLE_DATA[0];items.push(`<strong>${esc(String(top[0]))}</strong> has highest count: <strong>${fmtN(top[1])}</strong>.`);}
    const d5=addYears(5),d10=addYears(10);
    const r5=active.filter(e=>e.rd<=d5).length,r10=active.filter(e=>e.rd<=d10).length;
    if(r5>0)items.push(`<strong>${fmtN(r5)}</strong> employees retire in 5 years.`);
    if(r10>0&&pool.length){const pct=(r10/pool.length*100).toFixed(1);items.push(`<strong>${pct}%</strong> of workforce retires in 10 years.`);}
    const avg=pool.length?pool.reduce((s,e)=>s+e.age,0)/pool.length:0;
    if(avg>0)items.push(`Average age is <strong>${avg.toFixed(1)} years</strong>.`);
    const filterParts=getAppliedFiltersText();
    if(filterParts.length)items.push(`<strong>${filterParts.length}</strong> filter(s) active — showing ${fmtN(pool.length)} employees.`);
    if(FS_LOCAL.clickedType)items.push(`Detail filtered by <strong>${FS_LOCAL.clickedType}</strong>: ${FS_LOCAL.clickedValues.map(v=>`<em>${esc(v)}</em>`).join(', ')}.`);
    if(FS_LOCAL.timePeriod){const lbl={thisYear:'This Year',nextYear:'Next Year',next5y:'Next 5 Years'}[FS_LOCAL.timePeriod];items.push(`Time period filter: <strong>${lbl}</strong>.`);}
    document.getElementById('fs-insights').innerHTML=items.map(t=>`<div class="fs-insight-item">${t}</div>`).join('');
}

function refreshFS() {
    const basePool=getFSBasePool();
    const displayPool=getFSDisplayPool();
    if(FS_CH){FS_CH.destroy();FS_CH=null;}
    updateFSAppliedFilters();
    updateFSStats(displayPool);
    const ctx=document.getElementById('ca-fs');
    FS_TABLE_KEY='Item';

    if(FS_TYPE==='dept'){
        const m={};basePool.forEach(e=>m[e.dept]=(m[e.dept]||0)+1);const s=Object.entries(m).sort((a,b)=>b[1]-a[1]);
        const active=FS_LOCAL.clickedValues;
        const bgColors=s.map((d,i)=>{if(FS_LOCAL.clickedType==='dept'&&active.length)return active.includes(d[0])?PAL[i%PAL.length]+'ff':PAL[i%PAL.length]+'33';return PAL[i%PAL.length]+'bb';});
        FS_CH=new Chart(ctx,{type:'bar',data:{labels:s.map(x=>x[0]),datasets:[{label:'Count',data:s.map(x=>x[1]),backgroundColor:bgColors,borderRadius:4}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true},x:{ticks:{maxRotation:45}}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('dept',s[els[0].index][0]);}}});
        FS_TABLE_KEY='Department';
    } else if(FS_TYPE==='skill'){
        const m={};basePool.forEach(e=>m[e.skill]=(m[e.skill]||0)+1);const s=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,20);
        const bgColors=s.map(x=>(FS_LOCAL.clickedType==='skill'&&FS_LOCAL.clickedValues.length)?(FS_LOCAL.clickedValues.includes(x[0])?'#7c3aedff':'#7c3aed22'):'#7c3aed99');
        FS_CH=new Chart(ctx,{type:'bar',data:{labels:s.map(x=>x[0]),datasets:[{label:'Count',data:s.map(x=>x[1]),backgroundColor:bgColors,borderRadius:4}]},options:{...CBASE,indexAxis:'y',plugins:{...CBASE.plugins,legend:{display:false}},scales:{x:{beginAtZero:true}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('skill',s[els[0].index][0]);}}});
        FS_TABLE_KEY='Skill';
    } else if(FS_TYPE==='fy'){
        const m={};basePool.forEach(e=>m[e.fy]=(m[e.fy]||0)+1);const s=Object.entries(m).sort((a,b)=>getFYYear(a[0])-getFYYear(b[0]));
        const bgColors=s.map((x,i)=>(FS_LOCAL.clickedType==='fy'&&FS_LOCAL.clickedValues.length)?(FS_LOCAL.clickedValues.includes(x[0])?PAL[i%PAL.length]+'ff':PAL[i%PAL.length]+'33'):PAL[i%PAL.length]+'bb');
        FS_CH=new Chart(ctx,{type:'bar',data:{labels:s.map(x=>x[0]),datasets:[{label:'Retirements',data:s.map(x=>x[1]),backgroundColor:bgColors,borderRadius:4}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('fy',s[els[0].index][0]);}}});
        FS_TABLE_KEY='Financial Year';
    } else if(FS_TYPE==='month'){
        const mm=new Map();basePool.forEach(e=>{if(!mm.has(e.rdMS))mm.set(e.rdMS,{lbl:e.rdMonth,cnt:0});mm.get(e.rdMS).cnt++;});
        const s=[...mm.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
        const nowL=TODAY;
        const barColors=s.map(x=>{const yr=parseInt(x[0].split('-')[0]),mo=parseInt(x[0].split('-')[1])-1;const d=new Date(yr,mo,1);let base;if(d<new Date(nowL.getFullYear(),nowL.getMonth(),1))base='#94a3b8';else if(yr===nowL.getFullYear())base='#d97706';else if(yr===nowL.getFullYear()+1)base='#2563eb';else base='#065f46';if(FS_LOCAL.clickedType==='month'&&FS_LOCAL.clickedValues.length)return FS_LOCAL.clickedValues.includes(x[1].lbl)?base+'ff':base+'33';return base+'bb';});
        FS_CH=new Chart(ctx,{type:'bar',data:{labels:s.map(x=>x[1].lbl),datasets:[{label:'Retirees',data:s.map(x=>x[1].cnt),backgroundColor:barColors,borderColor:barColors.map(c=>c.replace(/bb$|33$|ff$/,'ff')),borderWidth:1,borderRadius:3}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true},x:{ticks:{maxRotation:45}}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('month',s[els[0].index][1].lbl);}}});
        FS_TABLE_KEY='Month';FS_TABLE_DATA=s.map(x=>[x[1].lbl,x[1].cnt]);
    } else if(FS_TYPE==='age-dist'){
        const bands=[{lo:20,hi:30,lbl:'20-30',c:'#6366f1'},{lo:31,hi:40,lbl:'31-40',c:'#8b5cf6'},{lo:41,hi:50,lbl:'41-50',c:'#3b82f6'},{lo:51,hi:55,lbl:'51-55',c:'#f59e0b'},{lo:56,hi:59,lbl:'56-59',c:'#ef4444'},{lo:60,hi:999,lbl:'60+',c:'#991b1b'}];
        const cnts=bands.map(b=>({...b,cnt:basePool.filter(e=>e.age>=b.lo&&e.age<=b.hi).length}));
        const bgColors=cnts.map(b=>(FS_LOCAL.clickedType==='ageBand'&&FS_LOCAL.clickedValues.length)?(FS_LOCAL.clickedValues.includes(b.lbl)?b.c+'ff':b.c+'33'):b.c+'bb');
        FS_CH=new Chart(ctx,{type:'bar',data:{labels:cnts.map(b=>b.lbl),datasets:[{label:'Employees',data:cnts.map(b=>b.cnt),backgroundColor:bgColors,borderColor:cnts.map(b=>b.c),borderRadius:5}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('ageBand',cnts[els[0].index].lbl);}}});
        FS_TABLE_KEY='Age Band';FS_TABLE_DATA=cnts.map(b=>[b.lbl+' yrs',b.cnt]);
    } else if(FS_TYPE==='grade'){
        const m={};basePool.forEach(e=>m[e.grade]=(m[e.grade]||0)+1);
        const s=Object.entries(m).sort((a,b)=>{const na=parseFloat((a[0].match(/\d+(\.\d+)?/)||[0])[0]);const nb=parseFloat((b[0].match(/\d+(\.\d+)?/)||[0])[0]);if(!isNaN(na)&&!isNaN(nb))return nb-na;return b[0].localeCompare(a[0]);});
        const bgColors=s.map(x=>(FS_LOCAL.clickedType==='grade'&&FS_LOCAL.clickedValues.length)?(FS_LOCAL.clickedValues.includes(x[0])?'#d97706cc':'#d9770622'):'#d9770699');
        FS_CH=new Chart(ctx,{type:'bar',data:{labels:s.map(x=>x[0]),datasets:[{label:'Count',data:s.map(x=>x[1]),backgroundColor:bgColors,borderRadius:4}]},options:{...CBASE,plugins:{...CBASE.plugins,legend:{display:false}},scales:{y:{beginAtZero:true}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('grade',s[els[0].index][0]);}}});
        FS_TABLE_KEY='Grade';
    } else if(FS_TYPE==='desig'){
        const m={};basePool.forEach(e=>m[e.desig]=(m[e.desig]||0)+1);const s=Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,15);
        const bgColors=s.map((x,i)=>(FS_LOCAL.clickedType==='desig'&&FS_LOCAL.clickedValues.length)?(FS_LOCAL.clickedValues.includes(x[0])?PAL[i%PAL.length]:PAL[i%PAL.length]+'44'):PAL[i%PAL.length]);
        FS_CH=new Chart(ctx,{type:'doughnut',data:{labels:s.map(x=>x[0]),datasets:[{data:s.map(x=>x[1]),backgroundColor:bgColors,borderWidth:2,borderColor:DARK?'#1a1d2e':'#fff'}]},options:{...CBASE,cutout:'50%',plugins:{...CBASE.plugins,legend:{position:'right'}},onClick:(ev,els)=>{if(els.length)fsCategoryClick('desig',s[els[0].index][0]);}}});
        FS_TABLE_KEY='Designation';
    }

    if(FS_TYPE!=='month'&&FS_TYPE!=='age-dist'){updateFSTable(displayPool);}
    else{document.getElementById('fs-thead').innerHTML=`<tr><th>${FS_TABLE_KEY}</th><th style="text-align:right">Count</th></tr>`;document.getElementById('fs-tbody').innerHTML=FS_TABLE_DATA.map(([k,v])=>`<tr><td class="cn">${esc(String(k))}</td><td style="text-align:right;font-family:var(--font-mono);font-weight:700">${fmtN(v)}</td></tr>`).join('');}
    updateFSInsights(displayPool);
    updateFSAppliedFilters();
}

function _buildFSTableDataFromPool(type,pool){
    FS_TABLE_KEY='Item';
    if(type==='dept'){const m={};pool.forEach(e=>m[e.dept]=(m[e.dept]||0)+1);FS_TABLE_KEY='Department';return Object.entries(m).sort((a,b)=>b[1]-a[1]);}
    else if(type==='skill'){const m={};pool.forEach(e=>m[e.skill]=(m[e.skill]||0)+1);FS_TABLE_KEY='Skill';return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,20);}
    else if(type==='fy'){const m={};pool.forEach(e=>m[e.fy]=(m[e.fy]||0)+1);FS_TABLE_KEY='Financial Year';return Object.entries(m).sort((a,b)=>getFYYear(a[0])-getFYYear(b[0]));}
    else if(type==='month'){const mm=new Map();pool.forEach(e=>{if(!mm.has(e.rdMS))mm.set(e.rdMS,{lbl:e.rdMonth,cnt:0});mm.get(e.rdMS).cnt++;});FS_TABLE_KEY='Month';return[...mm.entries()].sort((a,b)=>a[0].localeCompare(b[0])).map(x=>[x[1].lbl,x[1].cnt]);}
    else if(type==='age-dist'){const bands=[{lo:20,hi:30,lbl:'20-30'},{lo:31,hi:40,lbl:'31-40'},{lo:41,hi:50,lbl:'41-50'},{lo:51,hi:55,lbl:'51-55'},{lo:56,hi:59,lbl:'56-59'},{lo:60,hi:999,lbl:'60+'}];FS_TABLE_KEY='Age Band';return bands.map(b=>[b.lbl+' yrs',pool.filter(e=>e.age>=b.lo&&e.age<=b.hi).length]);}
    else if(type==='grade'){const m={};pool.forEach(e=>m[e.grade]=(m[e.grade]||0)+1);FS_TABLE_KEY='Grade';return Object.entries(m).sort((a,b)=>{const na=parseFloat((a[0].match(/\d+(\.\d+)?/)||[0])[0]);const nb=parseFloat((b[0].match(/\d+(\.\d+)?/)||[0])[0]);if(!isNaN(na)&&!isNaN(nb))return nb-na;return b[0].localeCompare(a[0]);});}
    else if(type==='desig'){const m={};pool.forEach(e=>m[e.desig]=(m[e.desig]||0)+1);FS_TABLE_KEY='Designation';return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,15);}
    return[];
}

// ===================== EXPORT =====================
function toggleExp(){document.getElementById('exp-dd').classList.toggle('open');}
function doExp(type){const data=getTableData();if(!data.length){alert('No data');return;}_exportEmpData(data,type,'NMDC_Employees');document.getElementById('exp-dd').classList.remove('open');}
function doExpAll(type){const data=[...DATA,...PAST];if(!data.length){alert('No data');return;}_exportEmpData(data,type,'NMDC_AllEmployees');document.getElementById('exp-dd').classList.remove('open');}
function _exportEmpData(data,type,fname){
    const hdrs=['Name','EmpID','Grade','Desig','Skill','Dept','SubDept','Deposit','DOB','RetDate','FY','Gender','Status'];
    const rows=data.map(e=>[e.name,e.empid,e.grade,e.desig,e.skill,e.dept,e.subdept,e.deposit,fmtDate(e.dob),fmtDate(e.rd),e.fy,e.gender,getStatus(e)]);
    if(type==='xlsx'){const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Data');XLSX.writeFile(wb,fname+'.xlsx');showToast('Excel exported!','fa-file-excel');}
    else{window.print();showToast('Sent to printer','fa-print');}
}

function openGlobalExport(){document.getElementById('gexp-ov').classList.add('show');}
function closeGexpOv(e){if(e.target.id==='gexp-ov')document.getElementById('gexp-ov').classList.remove('show');}
function doGexp(type){
    document.getElementById('gexp-ov').classList.remove('show');
    if(type==='view-xlsx')doExp('xlsx');
    else if(type==='view-pdf')doExp('pdf');
    else if(type==='report-xlsx')exportReportXLSX();
    else if(type==='report-pdf')exportReportPDF();
    else if(type==='all-xlsx')doExpAll('xlsx');
}

// ===================== SECTION EXPORT =====================
function exportSectionXLSX(sec){
    const wb=XLSX.utils.book_new();
    if(sec==='skill'){const rows=[['Skill','Total','1Y','5Y','10Y'],...Object.entries(_lastSkillMap).sort((a,b)=>b[1].t-a[1].t).map(([k,v])=>[k,v.t,v.y1,v.y5,v.y10])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Skill');}
    else if(sec==='dept'){const rows=[['Deposit','Category','Dept','Sub-Dept','Total','1Y','5Y','10Y'],...Object.values(_lastDeptMap).map(v=>[v.dep,v.cat,v.dept,v.sub||'N/A',v.t,v.y1,v.y5,v.y10])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Dept');}
    else if(sec==='deposit'){const rows=[['Deposit','Total','1Y','5Y','10Y'],...Object.entries(_lastDepMap).map(([k,v])=>[k,v.t,v.y1,v.y5,v.y10])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Deposit');}
    else if(sec==='grade'){const rows=[['Grade','Total','1Y','5Y','10Y'],...Object.entries(_lastGradeMap).map(([k,v])=>[k,v.t,v.y1,v.y5,v.y10])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Grade');}
    else if(sec==='fy'){let cum=0;const rows=[['FY','Retirements','Cumulative','% Workforce'],..._lastFySorted.map(([fy,cnt])=>{cum+=cnt;const pct=((cnt/(_lastFySorted.reduce((s,x)=>s+x[1],0)))*100).toFixed(1);return[fy,cnt,cum,pct+'%'];})];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'FY Forecast');}
    else if(sec==='skillgap'){const rows=[['Skill','Headcount','5Y Loss','10Y Loss','Impact %','Risk'],..._lastSkillGap.map(r=>[r.k,r.v.t,r.v.y5,r.v.y10,r.p+'%',r.lv.toUpperCase()])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Skill Gap');}
    else if(sec==='shortfall'){const rows=[['Dept','Sub-Dept','Strength','10Y Ret','Impact%','Risk'],..._lastShortfall.map(r=>[r.dept,r.sub||'N/A',r.s,r.r10,r.p+'%',r.lv.toUpperCase()])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Shortfall');}
    else if(sec==='desig'){const rows=[['Designation','Total','10Y Retirements'],...Object.entries(_lastDgMap).sort((a,b)=>b[1].t-a[1].t).map(([k,v])=>[k,v.t,v.r10])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Designation');}
    XLSX.writeFile(wb,`NMDC_${sec}.xlsx`);showToast(`${sec} exported!`,'fa-file-excel');
}

// ===================== REPORT SECTION CHECKBOX LOGIC =====================
function toggleRptSectionAll(sec,cb){const tbody=document.getElementById(`rs-${sec}-body`);if(!tbody)return;tbody.querySelectorAll('input[type=checkbox]').forEach(c=>{c.checked=cb.checked;const row=c.closest('tr');if(row)row.classList.toggle('row-selected',cb.checked);});const hdrCB=document.getElementById(`rsa-${sec}-h`);const barCB=document.getElementById(`rsa-${sec}`);if(hdrCB)hdrCB.checked=cb.checked;if(barCB)barCB.checked=cb.checked;updateRptSelBar(sec);}
function onRptRowCBChange(cb,sec){const row=cb.closest('tr');if(row)row.classList.toggle('row-selected',cb.checked);updateRptSelBar(sec);const tbody=document.getElementById(`rs-${sec}-body`);if(!tbody)return;const all=tbody.querySelectorAll('input[type=checkbox]');const checked=[...all].filter(c=>c.checked).length;const hdrCB=document.getElementById(`rsa-${sec}-h`);const barCB=document.getElementById(`rsa-${sec}`);if(hdrCB){hdrCB.indeterminate=checked>0&&checked<all.length;hdrCB.checked=checked===all.length&&all.length>0;}if(barCB){barCB.indeterminate=checked>0&&checked<all.length;barCB.checked=checked===all.length&&all.length>0;}}
function updateRptSelBar(sec){const tbody=document.getElementById(`rs-${sec}-body`);if(!tbody)return;const all=tbody.querySelectorAll('input[type=checkbox]');const checked=[...all].filter(c=>c.checked).length;const cntEl=document.getElementById(`rsc-${sec}`);const expBtn=document.getElementById(`rse-${sec}`);if(cntEl)cntEl.textContent=`${checked} selected`;if(expBtn)expBtn.style.display=checked>0?'inline-flex':'none';}
function exportRptSelection(sec){const tbody=document.getElementById(`rs-${sec}-body`);if(!tbody)return;const selected=[];tbody.querySelectorAll('tr').forEach(tr=>{const cb=tr.querySelector('input[type=checkbox]');if(!cb||!cb.checked)return;const cells=tr.querySelectorAll('td');const row=[...cells].slice(1).map(td=>td.textContent.trim());selected.push(row);});if(!selected.length){showToast('No rows selected','fa-exclamation-circle');return;}const headers={skill:['Skill / Trade','Total','1Y','5Y','10Y'],dept:['Deposit','Category','Department','Sub-Dept','Total','1Y','5Y','10Y'],deposit:['Deposit','Total','1Y','5Y','10Y'],grade:['Grade','Total','1Y','5Y','10Y'],desig:['Designation','Total','10Y Retirements'],fy:['Financial Year','Retirements','Cumulative','% Workforce','Trend'],skillgap:['Skill','Headcount','5Y Loss','10Y Loss','Impact %','Risk'],shortfall:['Department','Sub-Dept','Strength','10Y Ret','Impact %','Risk']};const hdrs=headers[sec]||['Item','Value'];const ws=XLSX.utils.aoa_to_sheet([hdrs,...selected]);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sec);XLSX.writeFile(wb,`NMDC_${sec}_selected.xlsx`);showToast(`${selected.length} rows exported!`,'fa-file-excel');}

// ===================== RENDER REPORT =====================
function renderReport() {
    const pool=RPT_FILTERED&&RPT_FILTERED.length?RPT_FILTERED:[...DATA,...PAST];
    const total=pool.length;
    document.getElementById('rpt-total-count').textContent=fmtN(total);
    const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const d1=addYears(1),d5=addYears(5),d10=addYears(10);
    const d3m=addMonths(3);
    const tm1=new Date(now.getFullYear(),now.getMonth(),1),tm2=new Date(now.getFullYear(),now.getMonth()+1,0);
    const nm1=new Date(now.getFullYear(),now.getMonth()+1,1),nm2=new Date(now.getFullYear(),now.getMonth()+2,0);
    const retiredCnt=pool.filter(e=>e.rd<now).length,activeCnt=active.length;
    const r1=active.filter(e=>e.rd<=d1).length,r5=active.filter(e=>e.rd<=d5).length,r10=active.filter(e=>e.rd<=d10).length;
    const r3m=active.filter(e=>e.rd<=d3m).length,rTM=pool.filter(e=>e.rd>=tm1&&e.rd<=tm2).length,rNM=pool.filter(e=>e.rd>=nm1&&e.rd<=nm2).length;
    const avgAge=total?pool.reduce((s,e)=>s+e.age,0)/total:0;

    const mkKPI=(num,lbl,sub,stripe,iBg,iClr,icon,onclick)=>`<div class="rpt-kpi-card" onclick="${onclick||''}" style="cursor:${onclick?'pointer':'default'}"><div class="rpt-kpi-stripe" style="background:${stripe}"></div><div class="rpt-kpi-icon" style="background:${iBg};color:${iClr}"><i class="fas ${icon}"></i></div><div class="rpt-kpi-num" style="color:${iClr}">${num}</div><div class="rpt-kpi-lbl">${lbl}</div>${sub?`<div class="rpt-kpi-sub">${sub}</div>`:''}</div>`;
    document.getElementById('rpt-kpi-workforce').innerHTML=
        mkKPI(fmtN(total),'Total Employees','All in dataset','linear-gradient(90deg,#1e3a8a,#2563eb)','#eff6ff','#1e3a8a','fa-users','')+
        mkKPI(fmtN(activeCnt),'Active Employees','Currently serving','linear-gradient(90deg,#065f46,#059669)','#ecfdf5','#065f46','fa-user-check','')+
        mkKPI(fmtN(retiredCnt),'Already Retired','Past retirement date','linear-gradient(90deg,#0c4a6e,#0284c7)','#f0f9ff','#0c4a6e','fa-user-minus',`openKPIDrill('retired')`)+
        mkKPI(avgAge>0?avgAge.toFixed(1)+' yrs':'—','Average Age','As of today','linear-gradient(90deg,#4c1d95,#7c3aed)','#f5f3ff','#4c1d95','fa-user-clock',`openKPIDrill('age')`);

    const retCards=[{n:rTM,l:'Retiring This Month',c:'#dc2626',type:'thisMonth'},{n:rNM,l:'Retiring Next Month',c:'#ea580c',type:'nextMonth'},{n:r3m,l:'Within 3 Months',c:'#d97706',type:'3m'},{n:r1,l:'Within 1 Year',c:'#ca8a04',type:'1y'},{n:r5,l:'Within 5 Years',c:'#65a30d',type:'5y'},{n:r10,l:'Within 10 Years',c:'#059669',type:'10y'},{n:retiredCnt,l:'Already Retired',c:'#0284c7',type:'retired'},{n:activeCnt,l:'Active Employees',c:'#065f46',type:'active'}];
    document.getElementById('std-ret-block').innerHTML=retCards.map(r=>`<div class="rpt-kpi-card" onclick="openKPIDrill('${r.type}')" style="cursor:pointer"><div class="rpt-kpi-stripe" style="background:${r.c}"></div><div class="rpt-kpi-num" style="color:${r.c};font-size:24px">${fmtN(r.n)}</div><div class="rpt-kpi-lbl">${r.l}</div><div class="rpt-kpi-sub">click for details →</div></div>`).join('');

    const rem5=active.filter(e=>e.rd>d5),rem10=active.filter(e=>e.rd>d10),rem3m=active.filter(e=>e.rd>d3m),rem1y=active.filter(e=>e.rd>d1);
    function avgA(arr){return arr.length?arr.reduce((s,e)=>s+e.age,0)/arr.length:0;}
    const ageProjs=[{n:avgAge.toFixed(1),l:'Today',c:'var(--ind)'},{n:rem3m.length?avgA(rem3m).toFixed(1):'—',l:'After 3 Months',c:'#0891b2'},{n:rem1y.length?avgA(rem1y).toFixed(1):'—',l:'After 1 Year',c:'#d97706'},{n:rem5.length?avgA(rem5).toFixed(1):'—',l:'After 5 Years',c:'#ea580c'},{n:rem10.length?avgA(rem10).toFixed(1):'—',l:'After 10 Years',c:'#dc2626'}];
    document.getElementById('age-metrics-row').innerHTML=ageProjs.map(a=>`<div class="age-metric"><div class="age-metric-num" style="color:${a.c}">${a.n} yrs</div><div class="age-metric-lbl">Avg Age ${a.l}</div></div>`).join('');
    genNarratives(pool,d1,d5,d10);

    const nT=n=>`<td class="num num-total">${fmtN(n)}</td>`,n1=n=>`<td class="num num-1y">${fmtN(n)}</td>`,n5=n=>`<td class="num num-5y">${fmtN(n)}</td>`,n10=n=>`<td class="num num-10y">${fmtN(n)}</td>`;

    const skillMap={};pool.forEach(e=>{const k=e.skill||'N/A';if(!skillMap[k])skillMap[k]={t:0,y1:0,y5:0,y10:0};skillMap[k].t++;if(inH(e,d1))skillMap[k].y1++;if(inH(e,d5))skillMap[k].y5++;if(inH(e,d10))skillMap[k].y10++;});
    _lastSkillMap=skillMap;let st=0,sy1=0,sy5=0,sy10=0;
    document.getElementById('rs-skill-body').innerHTML=Object.entries(skillMap).sort((a,b)=>b[1].t-a[1].t).map(([k,v])=>{st+=v.t;sy1+=v.y1;sy5+=v.y5;sy10+=v.y10;return`<tr onclick="showDrillBy('skill','${esc(k)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'skill')"></td><td><strong>${esc(k)}</strong></td>${nT(v.t)}${n1(v.y1)}${n5(v.y5)}${n10(v.y10)}</tr>`;}).join('')+`<tr class="rpt-grand"><td colspan="2">GRAND TOTAL</td>${nT(st)}${n1(sy1)}${n5(sy5)}${n10(sy10)}</tr>`;
    updateRptSelBar('skill');

    const deptMap={};pool.forEach(e=>{const key=`${e.deposit}||${e.category}||${e.dept}||${e.subdept}`;if(!deptMap[key])deptMap[key]={dep:e.deposit,cat:e.category,dept:e.dept,sub:e.subdept,t:0,y1:0,y5:0,y10:0};deptMap[key].t++;if(inH(e,d1))deptMap[key].y1++;if(inH(e,d5))deptMap[key].y5++;if(inH(e,d10))deptMap[key].y10++;});
    _lastDeptMap=deptMap;let dt=0,dy1=0,dy5=0,dy10=0;
    document.getElementById('rs-dept-body').innerHTML=Object.values(deptMap).sort((a,b)=>a.dep.localeCompare(b.dep)||a.dept.localeCompare(b.dept)||a.sub.localeCompare(b.sub)).map(v=>{dt+=v.t;dy1+=v.y1;dy5+=v.y5;dy10+=v.y10;return`<tr onclick="showDrillBy('dept','${esc(v.dept)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'dept')"></td><td>${esc(v.dep)}</td><td>${esc(v.cat)}</td><td><strong>${esc(v.dept)}</strong></td><td style="color:var(--tx2)">${esc(v.sub)}</td>${nT(v.t)}${n1(v.y1)}${n5(v.y5)}${n10(v.y10)}</tr>`;}).join('')+`<tr class="rpt-grand"><td colspan="5">GRAND TOTAL</td>${nT(dt)}${n1(dy1)}${n5(dy5)}${n10(dy10)}</tr>`;
    updateRptSelBar('dept');

    const depMap={};pool.forEach(e=>{const k=e.deposit||'N/A';if(!depMap[k])depMap[k]={t:0,y1:0,y5:0,y10:0};depMap[k].t++;if(inH(e,d1))depMap[k].y1++;if(inH(e,d5))depMap[k].y5++;if(inH(e,d10))depMap[k].y10++;});
    _lastDepMap=depMap;let dpt=0,dpy1=0,dpy5=0,dpy10=0;
    document.getElementById('rs-deposit-body').innerHTML=Object.entries(depMap).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>{dpt+=v.t;dpy1+=v.y1;dpy5+=v.y5;dpy10+=v.y10;return`<tr><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'deposit')"></td><td><strong>${esc(k)}</strong></td>${nT(v.t)}${n1(v.y1)}${n5(v.y5)}${n10(v.y10)}</tr>`;}).join('')+`<tr class="rpt-grand"><td colspan="2">GRAND TOTAL</td>${nT(dpt)}${n1(dpy1)}${n5(dpy5)}${n10(dpy10)}</tr>`;
    updateRptSelBar('deposit');

    const gradeMap={};pool.forEach(e=>{const k=e.grade||'N/A';if(!gradeMap[k])gradeMap[k]={t:0,y1:0,y5:0,y10:0};gradeMap[k].t++;if(inH(e,d1))gradeMap[k].y1++;if(inH(e,d5))gradeMap[k].y5++;if(inH(e,d10))gradeMap[k].y10++;});
    _lastGradeMap=gradeMap;let gt=0,gy1=0,gy5=0,gy10=0;
    document.getElementById('rs-grade-body').innerHTML=Object.entries(gradeMap).sort((a,b)=>{const numA=parseFloat((a[0].match(/\d+(\.\d+)?/)||[0])[0]);const numB=parseFloat((b[0].match(/\d+(\.\d+)?/)||[0])[0]);if(!isNaN(numA)&&!isNaN(numB))return numB-numA;return b[0].localeCompare(a[0]);}).map(([k,v])=>{gt+=v.t;gy1+=v.y1;gy5+=v.y5;gy10+=v.y10;return`<tr onclick="showDrillBy('grade','${esc(k)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'grade')"></td><td>${esc(k)}</td>${nT(v.t)}${n1(v.y1)}${n5(v.y5)}${n10(v.y10)}</tr>`;}).join('')+`<tr class="rpt-grand"><td colspan="2">GRAND TOTAL</td>${nT(gt)}${n1(gy1)}${n5(gy5)}${n10(gy10)}</tr>`;
    updateRptSelBar('grade');

    const dgMap={};pool.forEach(e=>{const k=e.desig||'N/A';if(!dgMap[k])dgMap[k]={t:0,r10:0};dgMap[k].t++;if(inH(e,d10))dgMap[k].r10++;});
    _lastDgMap=dgMap;let dgt=0,dgr=0;
    document.getElementById('rs-desig-body').innerHTML=Object.entries(dgMap).sort((a,b)=>b[1].t-a[1].t).map(([k,v])=>{dgt+=v.t;dgr+=v.r10;return`<tr onclick="showDrillBy('desig','${esc(k)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'desig')"></td><td>${esc(k)}</td>${nT(v.t)}<td class="num num-1y">${fmtN(v.r10)}</td></tr>`;}).join('')+`<tr class="rpt-grand"><td colspan="2">GRAND TOTAL</td>${nT(dgt)}<td class="num">${fmtN(dgr)}</td></tr>`;
    updateRptSelBar('desig');

    const tlRows=[{l:'This Month',n:rTM,c:'#dc2626'},{l:'Next Month',n:rNM,c:'#ea580c'},{l:'Within 3 Months',n:r3m,c:'#d97706'},{l:'Within 1 Year',n:r1,c:'#ca8a04'},{l:'Within 5 Years',n:r5,c:'#65a30d'},{l:'Within 10 Years',n:r10,c:'#059669'}];
    document.getElementById('rs-timeline-body').innerHTML=tlRows.map(r=>{const p=total>0?(r.n/total*100).toFixed(1):'0';const bw=total>0?Math.round(r.n/total*100):0;return`<tr><td><strong>${r.l}</strong></td><td class="num" style="color:${r.c};font-weight:800">${fmtN(r.n)}</td><td class="num"><div style="display:flex;align-items:center;gap:7px;justify-content:flex-end"><div style="width:75px;height:7px;background:var(--s3);border-radius:3px;overflow:hidden"><div style="width:${bw}%;height:100%;background:${r.c};border-radius:3px"></div></div><span style="font-weight:700;color:${r.c}">${p}%</span></div></td></tr>`;}).join('');

    const fyMap={};pool.filter(e=>e.rd>=now).forEach(e=>fyMap[e.fy]=(fyMap[e.fy]||0)+1);
    const fySorted=Object.entries(fyMap).sort((a,b)=>getFYYear(a[0])-getFYYear(b[0]));
    _lastFySorted=fySorted;let cum=0;
    document.getElementById('rs-fy-body').innerHTML=fySorted.map(([fy,cnt],i)=>{cum+=cnt;const pct=total>0?(cnt/total*100).toFixed(1):'0';const trend=i===0?'—':cnt>fySorted[i-1][1]?'<span style="color:#dc2626">↑ Higher</span>':'<span style="color:#059669">↓ Lower</span>';return`<tr onclick="showDrillBy('fy','${esc(fy)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'fy')"></td><td><span class="fy-chip">${esc(fy)}</span></td><td class="num num-1y">${fmtN(cnt)}</td><td class="num num-total">${fmtN(cum)}</td><td class="num"><span style="font-weight:700;color:#2563eb">${pct}%</span></td><td>${trend}</td></tr>`;}).join('');
    updateRptSelBar('fy');

    const ageBands=[{lo:20,hi:30,l:'20-30'},{lo:31,hi:40,l:'31-40'},{lo:41,hi:50,l:'41-50'},{lo:51,hi:55,l:'51-55'},{lo:56,hi:59,l:'56-59'},{lo:60,hi:999,l:'60+'}];
    document.getElementById('rs-age-body').innerHTML=ageBands.map(b=>{const emps=pool.filter(e=>e.age>=b.lo&&e.age<=b.hi);const r10c=emps.filter(e=>inH(e,d10)).length;const pct=total>0?(emps.length/total*100).toFixed(1):'0';return`<tr onclick="showDrillBy('ageBand','${b.l}')"><td><strong>${b.l} years</strong></td>${nT(emps.length)}<td class="num"><span style="font-weight:700;color:#2563eb">${pct}%</span></td>${n10(r10c)}</tr>`;}).join('');

    const sgItems=Object.entries(skillMap).map(([k,v])=>{const p5=v.t>0?Math.round(v.y5/v.t*100):0;const lv=riskLvl(p5);return{k,v,p5,lv};}).sort((a,b)=>b.p5-a.p5);
    _lastSkillGap=sgItems;
    document.getElementById('rs-skillgap-body').innerHTML=sgItems.map(({k,v,p5,lv})=>`<tr onclick="showDrillBy('skill','${esc(k)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'skillgap')"></td><td><strong>${esc(k)}</strong></td>${nT(v.t)}${n5(v.y5)}${n10(v.y10)}<td class="num"><div style="display:flex;align-items:center;gap:5px;justify-content:flex-end"><div style="width:55px;height:7px;background:var(--s3);border-radius:3px;overflow:hidden"><div style="width:${Math.min(p5,100)}%;height:100%;background:${p5>=40?'#dc2626':p5>=25?'#d97706':'#2563eb'};border-radius:3px"></div></div><span style="font-weight:800;color:${p5>=40?'#dc2626':p5>=25?'#d97706':'#2563eb'}">${p5}%</span></div></td><td>${riskBadge(lv)}</td></tr>`).join('');
    updateRptSelBar('skillgap');

    const riskItems=[];
    Object.entries(skillMap).forEach(([k,v])=>{if(v.y10>0){const p=total>0?Math.round(v.y10/total*100):0;riskItems.push({label:k+' (Skill)',cnt:v.y10,p,lv:riskLvl(p)});}});
    const dMap2={};pool.forEach(e=>{const k=`${e.dept}||${e.subdept}`;if(!dMap2[k])dMap2[k]={dept:e.dept,sub:e.subdept,t:0,y10:0};dMap2[k].t++;if(inH(e,d10))dMap2[k].y10++;});
    Object.values(dMap2).forEach(v=>{if(v.y10>0){const p=total>0?Math.round(v.y10/total*100):0;const label=v.dept+(v.sub&&v.sub!=='N/A'?' / '+v.sub:'')+'(Dept)';riskItems.push({label,cnt:v.y10,p,lv:riskLvl(p)});}});
    riskItems.sort((a,b)=>b.p-a.p);
    document.getElementById('rs-risk-body').innerHTML=riskItems.slice(0,20).map(r=>`<tr><td><strong>${esc(r.label)}</strong></td><td class="num num-1y">${fmtN(r.cnt)}</td><td class="num"><div style="display:flex;align-items:center;gap:5px;justify-content:flex-end"><div style="width:65px;height:7px;background:var(--s3);border-radius:3px;overflow:hidden"><div style="width:${Math.min(r.p,100)}%;height:100%;background:${r.p>=40?'#dc2626':r.p>=25?'#d97706':'#2563eb'};border-radius:3px"></div></div><span style="font-weight:800;color:${r.p>=40?'#dc2626':r.p>=25?'#d97706':'#2563eb'}">${r.p}%</span></div></td><td>${riskBadge(r.lv)}</td></tr>`).join('')||'<tr><td colspan="4" class="rpt-nodata">No risk data</td></tr>';

    const sfRows=Object.values(dMap2).map(v=>{const p=v.t>0?Math.round(v.y10/v.t*100):0;return{dept:v.dept,sub:v.sub,s:v.t,r10:v.y10,p,lv:riskLvl(p)};}).sort((a,b)=>b.p-a.p);
    _lastShortfall=sfRows;
    document.getElementById('rs-shortfall-body').innerHTML=(sfRows.length?sfRows.map(r=>`<tr onclick="showDrillBy('dept','${esc(r.dept)}')"><td class="cb-col" onclick="event.stopPropagation()"><input type="checkbox" onchange="onRptRowCBChange(this,'shortfall')"></td><td><strong>${esc(r.dept)}</strong></td><td style="color:var(--tx2)">${esc(r.sub)}</td>${nT(r.s)}<td class="num num-1y">${fmtN(r.r10)}</td><td class="num"><div style="display:flex;align-items:center;gap:5px;justify-content:flex-end"><div style="width:55px;height:7px;background:var(--s3);border-radius:3px;overflow:hidden"><div style="width:${Math.min(r.p,100)}%;height:100%;background:${r.p>=40?'#dc2626':r.p>=25?'#d97706':'#2563eb'};border-radius:3px"></div></div><span style="font-weight:800;color:${r.p>=40?'#dc2626':r.p>=25?'#d97706':'#2563eb'}">${r.p}%</span></div></td><td>${riskBadge(r.lv)}</td></tr>`).join(''):'<tr><td colspan="7" class="rpt-nodata">No data</td></tr>')+
        `<tr class="rpt-grand"><td colspan="3">GRAND TOTAL</td>${nT(total)}<td class="num">${fmtN(r10)}</td><td class="num">—</td><td>—</td></tr>`;
    updateRptSelBar('shortfall');
}

// ===================== NARRATIVES =====================
function genNarratives(pool,d1,d5,d10){
    if(!pool.length){document.getElementById('narrative-items').innerHTML='<div class="narrative-item" style="color:var(--tx3)">No data.</div>';return;}
    const items=[];const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const skillMap={};pool.forEach(e=>{const k=e.skill||'N/A';if(!skillMap[k])skillMap[k]={y1:0,y5:0,y10:0};if(inH(e,d1))skillMap[k].y1++;if(inH(e,d5))skillMap[k].y5++;if(inH(e,d10))skillMap[k].y10++;});
    const ts1=Object.entries(skillMap).sort((a,b)=>b[1].y1-a[1].y1)[0];
    if(ts1&&ts1[1].y1>0)items.push(`<strong>${fmtN(ts1[1].y1)}</strong> employees with <em>${esc(ts1[0])}</em> skill retire in 1 year — immediate succession planning needed.`);
    const deptMap={};pool.forEach(e=>{const k=e.dept||'N/A';if(!deptMap[k])deptMap[k]={y5:0};if(inH(e,d5))deptMap[k].y5++;});
    const td5=Object.entries(deptMap).sort((a,b)=>b[1].y5-a[1].y5)[0];
    if(td5&&td5[1].y5>0)items.push(`Dept <strong>${esc(td5[0])}</strong> loses <strong>${fmtN(td5[1].y5)}</strong> employees in 5 years.`);
    const gradeMap={};pool.forEach(e=>{const k=e.grade||'N/A';if(!gradeMap[k])gradeMap[k]={y10:0};if(inH(e,d10))gradeMap[k].y10++;});
    const tg=Object.entries(gradeMap).sort((a,b)=>b[1].y10-a[1].y10)[0];
    if(tg&&tg[1].y10>0)items.push(`Grade <strong>${esc(tg[0])}</strong> has highest retirement concentration — <strong>${fmtN(tg[1].y10)}</strong> in 10 years.`);
    const fyMap={};pool.filter(e=>e.rd>=now).forEach(e=>fyMap[e.fy]=(fyMap[e.fy]||0)+1);
    const tfy=Object.entries(fyMap).sort((a,b)=>b[1]-a[1])[0];
    if(tfy)items.push(`<strong>${esc(tfy[0])}</strong> is the peak retirement year with <strong>${fmtN(tfy[1])}</strong> retirements.`);
    const r10=active.filter(e=>e.rd<=d10).length;
    if(pool.length>0)items.push(`<strong>${(r10/pool.length*100).toFixed(1)}%</strong> of workforce retires in 10 years — strategic succession planning is critical.`);
    if(!items.length)items.push('No significant retirement risks in current dataset.');
    document.getElementById('narrative-items').innerHTML=items.map(t=>`<div class="narrative-item">${t}</div>`).join('');
}

// ===================== REPORT EXPORT =====================
function exportReportXLSX(){
    const pool=RPT_FILTERED&&RPT_FILTERED.length?RPT_FILTERED:[...DATA,...PAST];
    const wb=XLSX.utils.book_new();
    const sm=_lastSkillMap;
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Skill','Total','1Y','5Y','10Y'],...Object.entries(sm).sort((a,b)=>b[1].t-a[1].t).map(([k,v])=>[k,v.t,v.y1,v.y5,v.y10])]),'1-Skill');
    const dm=_lastDeptMap;
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Deposit','Category','Dept','Sub-Dept','Total','1Y','5Y','10Y'],...Object.values(dm).map(v=>[v.dep,v.cat,v.dept,v.sub||'N/A',v.t,v.y1,v.y5,v.y10])]),'2-Dept');
    const gm=_lastGradeMap;
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Grade','Total','1Y','5Y','10Y'],...Object.entries(gm).map(([k,v])=>[k,v.t,v.y1,v.y5,v.y10])]),'3-Grade');
    const fs=_lastFySorted;let cum=0;
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['FY','Retirements','Cumulative','% Workforce'],...fs.map(([fy,cnt])=>{cum+=cnt;return[fy,cnt,cum,((cnt/pool.length)*100).toFixed(1)+'%'];})]),'4-FY-Forecast');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Dept','Sub-Dept','Strength','10Y Ret','Impact%','Risk'],..._lastShortfall.map(r=>[r.dept,r.sub||'N/A',r.s,r.r10,r.p+'%',r.lv.toUpperCase()])]),'5-Shortfall');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Skill','Headcount','5Y Loss','10Y Loss','Impact%','Risk'],..._lastSkillGap.map(r=>[r.k,r.v.t,r.v.y5,r.v.y10,r.p+'%',r.lv.toUpperCase()])]),'6-SkillGap');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Name','EmpID','Grade','Desig','Skill','Dept','SubDept','Deposit','DOB','RetDate','FY','Gender','Age'],...pool.map(e=>[e.name,e.empid,e.grade,e.desig,e.skill,e.dept,e.subdept,e.deposit,fmtDate(e.dob),fmtDate(e.rd),e.fy,e.gender,e.age])]),'7-Raw');
    XLSX.writeFile(wb,'NMDC_Report.xlsx');showToast('Full report exported!','fa-file-excel');
}
function exportReportPDF(){
    document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
    document.querySelectorAll('.pg').forEach(p=>p.classList.remove('on'));
    document.querySelector('[data-p="report"]').classList.add('on');
    document.getElementById('pg-report').classList.add('on');
    setTimeout(()=>window.print(),200);
}
function exportFullReport(){exportReportXLSX();}

// ===================== CHART EXPORT MENUS =====================
function toggleFSExportMenu(e){e.stopPropagation();const menu=document.getElementById('fs-export-menu');const isOpen=menu.classList.contains('open');closeAllExportMenus();if(!isOpen)menu.classList.add('open');}
function toggleDrillExportMenu(e){e.stopPropagation();const menu=document.getElementById('drill-export-menu');const isOpen=menu.classList.contains('open');closeAllExportMenus();if(!isOpen)menu.classList.add('open');}
function toggleChartExportMenu(e,menuId){e.stopPropagation();const menu=document.getElementById(menuId);if(!menu)return;const isOpen=menu.classList.contains('open');closeAllExportMenus();if(!isOpen)menu.classList.add('open');}
function closeAllExportMenus(){document.querySelectorAll('.fs-export-menu,.chart-export-menu,#drill-export-menu').forEach(m=>m.classList.remove('open'));}
document.addEventListener('click',()=>closeAllExportMenus());

function getActiveFiltersText(){
    const parts=[];
    if(GF.dep.length)parts.push('Deposit: '+GF.dep.join(', '));if(GF.dept.length)parts.push('Dept: '+GF.dept.join(', '));if(GF.sub.length)parts.push('Sub-Dept: '+GF.sub.join(', '));if(GF.skill.length)parts.push('Skill: '+GF.skill.join(', '));if(GF.grade.length)parts.push('Grade: '+GF.grade.join(', '));if(GF.desig.length)parts.push('Desig: '+GF.desig.join(', '));if(GF.status)parts.push('Status: '+GF.status);if(GF.ry.length)parts.push('Ret. Year: '+GF.ry.join(', '));if(GF.rfy.length)parts.push('Ret. FY: '+GF.rfy.join(', '));if(GF.from)parts.push('From: '+GF.from);if(GF.to)parts.push('To: '+GF.to);
    return parts.length?parts.join(' | '):'No filters applied';
}

function exportFSChart(format){closeAllExportMenus();const canvas=document.getElementById('ca-fs');if(!canvas||!FS_CH){showToast('No chart to export','fa-exclamation-circle');return;}const title=document.getElementById('fs-title').textContent||'Chart';if(format==='png'||format==='jpeg'){_exportCanvasImage(canvas,format,title);}else if(format==='pdf'){_exportChartPDF(canvas,title,FS_TABLE_DATA,FS_TABLE_KEY);}}
function exportDrillChart(format){closeAllExportMenus();showToast('Drill view exports as table data only (no standalone chart)','fa-info-circle');}
function exportInlineChart(format,canvasId,title){closeAllExportMenus();const canvas=document.getElementById(canvasId);if(!canvas){showToast('Chart not available','fa-exclamation-circle');return;}if(format==='png'||format==='jpeg'){_exportCanvasImage(canvas,format,title);}else if(format==='pdf'){const chObj=CH[canvasId];const tableData=chObj?_extractChartTableData(chObj):[];_exportChartPDF(canvas,title,tableData,'Item');}}
function exportInlineChartExcel(type,title){closeAllExportMenus();const typeMap={'age-dist':'age-dist','dept':'dept','month':'month','desig':'desig','skill':'skill','fy':'fy','grade':'grade'};const fsType=typeMap[type];if(!fsType)return;const prevType=FS_TYPE;FS_TYPE=fsType;const prevTitle=document.getElementById('fs-title').textContent;document.getElementById('fs-title').textContent=title;const pool=applyGFToPool([...DATA,...PAST]);_buildFSTableDataFromPool(fsType,pool);exportFSExcel();FS_TYPE=prevType;document.getElementById('fs-title').textContent=prevTitle;}
function _extractChartTableData(chObj){try{const labels=chObj.data.labels||[];const data=chObj.data.datasets[0]?.data||[];return labels.map((l,i)=>[l,data[i]||0]);}catch(e){return[];}}

function exportFSExcel(){
    if(!FS_TABLE_DATA.length){alert('No data to export');return;}
    const pool=applyGFToPool([...DATA,...PAST]);const now=TODAY,active=pool.filter(e=>e.rd>=now);
    const tm1=new Date(now.getFullYear(),now.getMonth(),1),tm2=new Date(now.getFullYear(),now.getMonth()+1,0);
    const nm1=new Date(now.getFullYear(),now.getMonth()+1,1),nm2=new Date(now.getFullYear(),now.getMonth()+2,0);
    const d3m=addMonths(3),d1=addYears(1),d5=addYears(5),d10=addYears(10);
    function getEmpsByLabel(label){
        if(FS_TYPE==='dept')return pool.filter(e=>e.dept===label);if(FS_TYPE==='skill')return pool.filter(e=>e.skill===label||(e.skill==='N/A'&&e.desig===label));if(FS_TYPE==='desig')return pool.filter(e=>e.desig===label);if(FS_TYPE==='grade')return pool.filter(e=>e.grade===label);if(FS_TYPE==='fy')return pool.filter(e=>e.fy===label);if(FS_TYPE==='month')return pool.filter(e=>e.rdMonth===label);
        if(FS_TYPE==='age-dist'){const cleanLabel=label.replace(' yrs','');if(cleanLabel==='60+')return pool.filter(e=>e.age>=60);const[lo,hi]=cleanLabel.split('-').map(Number);return pool.filter(e=>e.age>=lo&&e.age<=hi);}return[];
    }
    const headers=[FS_TABLE_KEY,'Total','Retiring This Month','Retiring Next Month','Within 3 Months','Within 1 Year','Within 5 Years','Within 10 Years'];
    const dataRows=FS_TABLE_DATA.map(([label,total])=>{const emps=getEmpsByLabel(String(label));const act=emps.filter(e=>e.rd>=now);return[label,total,emps.filter(e=>e.rd>=tm1&&e.rd<=tm2).length,emps.filter(e=>e.rd>=nm1&&e.rd<=nm2).length,act.filter(e=>e.rd<=d3m).length,act.filter(e=>e.rd<=d1).length,act.filter(e=>e.rd<=d5).length,act.filter(e=>e.rd<=d10).length];});
    const filterParts=getActiveFiltersText();const filterRows=filterParts!=='No filters applied'?[[],['Applied Filters'],[filterParts],[]]:[[]];
    const totalsRow=['TOTAL',pool.length,pool.filter(e=>e.rd>=tm1&&e.rd<=tm2).length,pool.filter(e=>e.rd>=nm1&&e.rd<=nm2).length,active.filter(e=>e.rd<=d3m).length,active.filter(e=>e.rd<=d1).length,active.filter(e=>e.rd<=d5).length,active.filter(e=>e.rd<=d10).length];
    const aoa=[...filterRows,headers,...dataRows,[],totalsRow];
    const labelW=Math.max(FS_TABLE_KEY.length,...FS_TABLE_DATA.map(([k])=>String(k).length))+2;
    const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=[{wch:labelW},{wch:8},{wch:22},{wch:22},{wch:18},{wch:14},{wch:14},{wch:16}];
    const sheetName=(FS_TABLE_KEY+' Count').slice(0,31);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheetName);XLSX.writeFile(wb,`NMDC_${FS_TABLE_KEY.replace(/\s+/g,'_')}_Count.xlsx`);showToast(`${FS_TABLE_KEY} count exported!`,'fa-file-excel');
}

function _exportCanvasImage(canvas,format,title){
    const scale=3;const exportCanvas=document.createElement('canvas');exportCanvas.width=canvas.width*scale;exportCanvas.height=(canvas.height+80)*scale;const ctx=exportCanvas.getContext('2d');ctx.fillStyle=DARK?'#1a1d2e':'#ffffff';ctx.fillRect(0,0,exportCanvas.width,exportCanvas.height);ctx.fillStyle=DARK?'#e8eaf6':'#1A1D2E';ctx.font=`bold ${14*scale}px Source Sans 3, sans-serif`;ctx.fillText(title,12*scale,20*scale);ctx.fillStyle=DARK?'#6b75a0':'#828AAA';ctx.font=`${9*scale}px Source Sans 3, sans-serif`;const filterText='Filters: '+getActiveFiltersText();ctx.fillText(filterText.length>110?filterText.slice(0,110)+'…':filterText,12*scale,34*scale);ctx.fillText('Exported: '+new Date().toLocaleString('en-IN'),12*scale,46*scale);ctx.drawImage(canvas,0,58*scale,canvas.width*scale,canvas.height*scale);const mime=format==='jpeg'?'image/jpeg':'image/png';const quality=format==='jpeg'?0.92:1.0;const dataURL=exportCanvas.toDataURL(mime,quality);const a=document.createElement('a');a.href=dataURL;a.download=`NMDC_${title.replace(/\s+/g,'_')}.${format}`;a.click();showToast(`${format.toUpperCase()} exported!`,'fa-image');
}

function _exportChartPDF(canvas,title,tableData,tableKey){
    const filters=getActiveFiltersText();const timestamp=new Date().toLocaleString('en-IN');const isDark=DARK;const bg=isDark?'#1a1d2e':'#ffffff';const tx=isDark?'#e8eaf6':'#1A1D2E';const tx2=isDark?'#b0b8d8':'#3B4060';const bd=isDark?'#2e3352':'#D6DAE6';
    const exportCanvas=document.createElement('canvas');exportCanvas.width=canvas.width*2;exportCanvas.height=canvas.height*2;const ctx=exportCanvas.getContext('2d');ctx.fillStyle=bg;ctx.fillRect(0,0,exportCanvas.width,exportCanvas.height);ctx.drawImage(canvas,0,0,canvas.width*2,canvas.height*2);const chartImg=exportCanvas.toDataURL('image/png');
    const tableRows=tableData.map(([k,v])=>`<tr><td>${String(k)}</td><td style="text-align:right;font-weight:700">${Number(v).toLocaleString('en-IN')}</td></table>`).join('');
    const printWin=window.open('','_blank','width=900,height=700');if(!printWin){showToast('Allow popups for PDF export','fa-exclamation-circle');return;}
    printWin.document.write(`<!DOCTYPE html><html><head><title>${title}</title><style>body{font-family:'Source Sans 3',sans-serif;background:${bg};color:${tx};margin:0;padding:24px}.header{border-bottom:2px solid ${bd};padding-bottom:14px;margin-bottom:18px}h1{font-size:20px;margin:0 0 6px 0;font-family:Georgia,serif}.meta{font-size:11px;color:${tx2};line-height:1.6}.filters-box{background:${isDark?'#1e2235':'#f0f4ff'};border:1px solid ${bd};border-radius:6px;padding:10px 14px;margin:12px 0;font-size:11px}.filters-title{font-weight:700;text-transform:uppercase;font-size:9px;letter-spacing:.6px;color:${tx2};margin-bottom:4px}.chart-img{width:100%;max-width:100%;border:1px solid ${bd};border-radius:8px;display:block;margin:18px 0}.data-title{font-size:13px;font-weight:700;margin:18px 0 10px;color:${tx}}table{width:100%;border-collapse:collapse;font-size:12px}th{background:${isDark?'#242840':'#f0f4ff'};text-align:left;padding:8px 12px;border-bottom:2px solid ${bd};font-weight:700}td{padding:7px 12px;border-bottom:1px solid ${bd}}tr:nth-child(even) td{background:${isDark?'#1e2235':'#fafbfd'}}.footer{margin-top:24px;font-size:10px;color:${tx2};border-top:1px solid ${bd};padding-top:10px}@media print{body{padding:10px}button{display:none}}</style></head><body><div class="header"><h1>NMDC · HR Retirement Intelligence — ${title}</h1><div class="meta"><div><strong>Exported:</strong> ${timestamp}</div></div></div><div class="filters-box"><div class="filters-title">Applied Filters</div><div>${filters}</div></div><img class="chart-img" src="${chartImg}" alt="${title} chart">${tableRows?`<div class="data-title">Data Table — ${tableKey}</div><table><thead><tr><th>${tableKey}</th><th style="text-align:right">Count</th></tr></thead><tbody>${tableRows}</tbody></table>`:''}<div class="footer">NMDC HR Retirement Intelligence Dashboard · Generated on ${timestamp}</div><script>setTimeout(()=>{window.print();},400);<\/script></body></html>`);
    printWin.document.close();showToast('PDF print dialog opened!','fa-file-pdf');
}

// Hide shared GF wrapper initially (shown after data load)
document.addEventListener('DOMContentLoaded', () => {
    const wrapper = document.getElementById('shared-gf-wrapper');
    if (wrapper) wrapper.style.display = 'none';
});