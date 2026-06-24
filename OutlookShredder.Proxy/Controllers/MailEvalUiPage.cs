namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// The self-contained HTML/CSS/JS for the dev-only mail-classification eval UI, served by
/// MailEvalController.Ui() at GET /api/mail-eval/ui. No build step, no static-files host — one page,
/// vanilla JS, talking to the same-origin /api/mail-eval/* endpoints (exempt from WS2 auth).
/// Labeling loop: pick a row → read the body → choose the correct leaf → Save (sets LabeledBy).
/// Then Run Baseline scores the classifier over the human-labeled rows and renders the report.
/// </summary>
internal static class MailEvalUiPage
{
    public const string Html = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mail Classification Eval</title>
<style>
  :root { --bg:#1b1f27; --panel:#232a34; --panel2:#2b3340; --line:#374151; --fg:#e6e6e6; --mut:#9aa4b2;
          --accent:#3b82f6; --good:#22c55e; --warn:#f59e0b; --bad:#ef4444; }
  * { box-sizing:border-box; }
  body { margin:0; font:13px/1.45 -apple-system,Segoe UI,Roboto,sans-serif; background:var(--bg); color:var(--fg); }
  header { display:flex; align-items:center; gap:14px; padding:10px 16px; background:var(--panel); border-bottom:1px solid var(--line); flex-wrap:wrap; }
  header h1 { font-size:15px; margin:0; font-weight:600; }
  .chip { background:var(--panel2); border:1px solid var(--line); border-radius:14px; padding:3px 10px; font-size:12px; color:var(--mut); }
  .chip b { color:var(--fg); }
  button { background:var(--accent); color:#fff; border:0; border-radius:6px; padding:6px 12px; font-size:13px; cursor:pointer; }
  button.ghost { background:var(--panel2); color:var(--fg); border:1px solid var(--line); }
  button:disabled { opacity:.5; cursor:default; }
  main { display:grid; grid-template-columns:minmax(360px,1fr) minmax(420px,1.3fr); gap:0; height:calc(100vh - 53px); }
  .col { overflow:auto; padding:12px; }
  .col.left { border-right:1px solid var(--line); }
  .controls { display:flex; gap:8px; margin-bottom:10px; flex-wrap:wrap; }
  select, input[type=text] { background:var(--panel2); color:var(--fg); border:1px solid var(--line); border-radius:6px; padding:5px 8px; font-size:13px; }
  input[type=text] { flex:1; min-width:120px; }
  .row { padding:8px 10px; border:1px solid var(--line); border-radius:7px; margin-bottom:6px; cursor:pointer; background:var(--panel); }
  .row:hover { border-color:var(--accent); }
  .row.sel { border-color:var(--accent); background:var(--panel2); }
  .row .subj { font-weight:600; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .row .meta { color:var(--mut); font-size:11.5px; display:flex; gap:8px; margin-top:3px; align-items:center; }
  .badge { font-size:10.5px; padding:1px 6px; border-radius:10px; border:1px solid var(--line); }
  .badge.boot { color:var(--warn); border-color:var(--warn); }
  .badge.human { color:var(--good); border-color:var(--good); }
  .det h2 { font-size:14px; margin:0 0 4px; }
  .det .kv { color:var(--mut); font-size:12px; margin-bottom:2px; }
  .bodybox { background:#fff; color:#111; border-radius:6px; margin-top:10px; height:340px; width:100%; border:1px solid var(--line); }
  pre.bodytext { background:#fff; color:#111; border-radius:6px; padding:10px; margin-top:10px; height:340px; overflow:auto; white-space:pre-wrap; word-break:break-word; }
  .labelbar { display:flex; gap:8px; align-items:center; margin-top:12px; flex-wrap:wrap; }
  .labelbar select { flex:1; min-width:220px; }
  .hint { color:var(--mut); font-size:11.5px; margin-top:6px; }
  table { border-collapse:collapse; width:100%; font-size:12px; margin-top:8px; }
  th,td { border:1px solid var(--line); padding:4px 7px; text-align:left; }
  th { background:var(--panel2); position:sticky; top:0; }
  td.num { text-align:right; font-variant-numeric:tabular-nums; }
  .ok { color:var(--good); } .lo { color:var(--bad); } .mid { color:var(--warn); }
  #report { display:none; }
  .modal { position:fixed; inset:0; background:rgba(0,0,0,.55); display:none; align-items:flex-start; justify-content:center; overflow:auto; padding:24px; z-index:9; }
  .modal.open { display:flex; }
  .modal .card { background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:18px; max-width:980px; width:100%; }
  .modal h2 { margin:0 0 8px; }
  .sec { margin-top:16px; } .sec h3 { font-size:13px; color:var(--mut); margin:0 0 4px; text-transform:uppercase; letter-spacing:.04em; }
  .empty { color:var(--mut); padding:24px; text-align:center; }
  .toast { position:fixed; bottom:16px; left:50%; transform:translateX(-50%); background:var(--good); color:#06210f; padding:8px 16px; border-radius:8px; font-weight:600; opacity:0; transition:opacity .2s; }
  .toast.show { opacity:1; }
  .rulebox { margin-top:12px; border:1px solid var(--line); border-radius:7px; padding:8px 10px; background:var(--panel); }
  .rulebox summary { cursor:pointer; color:var(--accent); font-size:12.5px; }
  .ruleform { margin-top:8px; } .ruleform select, .ruleform input { margin:2px 0; }
</style>
</head>
<body>
<header>
  <h1>Mail Classification Eval</h1>
  <span class="chip">labeled <b id="cLabeled">–</b> / <b id="cTotal">–</b></span>
  <span class="chip">bootstrap <b id="cBoot">–</b></span>
  <span class="chip" id="cPct">–</span>
  <span style="flex:1"></span>
  <label class="chip">labeler <input type="text" id="labeler" value="angus" style="width:90px;flex:none"></label>
  <button class="ghost" onclick="reload()">Refresh</button>
  <button id="runBtn" onclick="runBaseline()">Run Baseline</button>
  <button class="ghost" onclick="openReport()">View Report</button>
  <button class="ghost" onclick="openImpact()">Rule Impact</button>
</header>

<main>
  <section class="col left">
    <div class="controls">
      <select id="fView" onchange="render()">
        <option value="unlabeled">Unlabeled (bootstrap)</option>
        <option value="labeled">Human-labeled</option>
        <option value="all">All</option>
      </select>
      <select id="fCat" onchange="render()"><option value="">All categories</option></select>
      <input type="text" id="fSearch" placeholder="search subject…" oninput="render()">
    </div>
    <div id="list"></div>
  </section>

  <section class="col det" id="detail">
    <div class="empty">Select a message on the left to read it and set its true category.</div>
  </section>
</main>

<div class="modal" id="reportModal" onclick="if(event.target===this)closeReport()">
  <div class="card" id="reportBody"></div>
</div>
<div class="modal" id="impactModal" onclick="if(event.target===this)closeImpact()">
  <div class="card" id="impactBody"></div>
</div>
<div class="toast" id="toast"></div>

<script>
const API = '/api/mail-eval';
let GOLD = { rows: [] }, LEAVES = [], CUR = null, IDX = -1;
const CRIT = ['Supplier/Order Confirmations','Supplier/Invoices and Bills','Supplier/Receipts','Supplier/RFQ Responses','Supplier/MTRs'];

const esc = s => (s||'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const isBoot = r => !r.labeledBy || r.labeledBy.toLowerCase() === 'bootstrap';

async function jget(p){ const r = await fetch(API+p); if(r.status===204) return null; if(!r.ok) throw new Error(p+' '+r.status); return r.json(); }
async function jpost(p,b){ const r = await fetch(API+p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)}); if(!r.ok) throw new Error(p+' '+r.status); return r.json(); }

function toast(msg){ const t=document.getElementById('toast'); t.textContent=msg; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),1400); }

async function reload(){
  GOLD = await jget('/golden');
  if(!LEAVES.length){ LEAVES = await jget('/leaves'); buildLeafPickers(); }
  // category filter options (with counts)
  const fc = document.getElementById('fCat'); const cur = fc.value;
  const cats = GOLD.byCategory || {};
  fc.innerHTML = '<option value="">All categories ('+GOLD.total+')</option>' +
    Object.keys(cats).map(k=>`<option value="${esc(k)}">${esc(k)} (${cats[k]})</option>`).join('');
  fc.value = cur;
  document.getElementById('cTotal').textContent = GOLD.total;
  document.getElementById('cLabeled').textContent = GOLD.humanLabeled;
  document.getElementById('cBoot').textContent = GOLD.bootstrap;
  const pct = GOLD.total ? Math.round(100*GOLD.humanLabeled/GOLD.total) : 0;
  document.getElementById('cPct').innerHTML = '<b>'+pct+'%</b> done';
  render();
}

function buildLeafPickers(){
  // ordered: critical leaves first, then the rest
  const ordered = [...CRIT.filter(c=>LEAVES.includes(c)), ...LEAVES.filter(l=>!CRIT.includes(l))];
  window._leafOpts = ordered;
}

function filtered(){
  const v = document.getElementById('fView').value;
  const cat = document.getElementById('fCat').value;
  const q = document.getElementById('fSearch').value.toLowerCase();
  return GOLD.rows.filter(r=>{
    if(v==='unlabeled' && !isBoot(r)) return false;
    if(v==='labeled' && isBoot(r)) return false;
    if(cat && r.goldenCategory!==cat) return false;
    if(q && !(r.subject||'').toLowerCase().includes(q)) return false;
    return true;
  });
}

function render(){
  const rows = filtered(); window._view = rows;
  const html = rows.map((r,i)=>`
    <div class="row ${CUR&&CUR.mailItemId===r.mailItemId?'sel':''}" data-i="${i}" onclick="pick(${i})">
      <div class="subj">${esc(r.subject)||'<i>(no subject)</i>'}</div>
      <div class="meta">
        <span class="badge ${isBoot(r)?'boot':'human'}">${isBoot(r)?'bootstrap':esc(r.labeledBy)}</span>
        <span>${esc(r.goldenCategory)||'—'}</span>
      </div>
    </div>`).join('');
  document.getElementById('list').innerHTML = html || '<div class="empty">No rows match this filter.</div>';
}

async function pick(i){
  const rows = window._view || []; const r = rows[i]; if(!r) return;
  IDX = i; CUR = r; render();
  const d = document.getElementById('detail');
  d.innerHTML = '<div class="empty">Loading…</div>';
  let det = null;
  try { det = await jget('/item/'+encodeURIComponent(r.mailItemId)); } catch(e){}
  const opts = (window._leafOpts||LEAVES).map(l=>`<option value="${esc(l)}" ${l===r.goldenCategory?'selected':''}>${esc(l)}</option>`).join('');
  const dom = (((det&&det.fromAddress)||r.fromAddress||'').split('@')[1]||'').trim();
  const sigOpt = s => ['SenderDomain','SenderAddress','Subject','Body'].map(x=>`<option ${x===s?'selected':''}>${x}</option>`).join('');
  const opOpt = ['Contains','Equals','Regex'].map(x=>`<option>${x}</option>`).join('');
  const body = det && det.isHtml && det.html
      ? `<iframe class="bodybox" sandbox srcdoc="${esc(det.html)}"></iframe>`
      : `<pre class="bodytext">${esc(det ? (det.bodyText||'(no body captured)') : '(could not load item detail)')}</pre>`;
  d.innerHTML = `
    <h2>${esc(r.subject)||'(no subject)'}</h2>
    <div class="kv">From: ${esc(det?det.fromAddress:r.fromAddress||'')} ${det&&det.fromName?'· '+esc(det.fromName):''}</div>
    <div class="kv">${det?('To: '+esc(det.toLine||'')+' · '+esc(det.receivedAt||'')):''}</div>
    <div class="kv">AI guess: <b>${esc(r.goldenCategory)}</b> ${det?('· conf '+(det.confidence!=null?Math.round(det.confidence*100)+'%':'?')+' · '+esc(det.aiProvider||'')):''}</div>
    ${body}
    <div class="labelbar">
      <select id="leafPick">${opts}</select>
      <button onclick="save()">Save (Enter)</button>
      <button class="ghost" onclick="addCategory()" title="Create a new category and select it">＋ New category</button>
      <button class="ghost" onclick="pick(Math.min((window._view.length-1), IDX+1))">Skip →</button>
    </div>
    <div class="hint">Critical leaves (over-sample these): ${CRIT.join(' · ')}</div>
    <details class="rulebox">
      <summary>＋ Make a deterministic rule from this</summary>
      <div class="ruleform">
        <div class="kv">When <select id="rSig">${sigOpt('SenderDomain')}</select>
          <select id="rOp">${opOpt}</select>
          <input type="text" id="rVal" value="${esc(dom)}" placeholder="value">
          → <select id="rCat">${opts}</select></div>
        <div class="labelbar">
          <input type="text" id="rName" placeholder="rule name" value="${esc(dom?dom+' → '+r.goldenCategory:'')}">
          <button onclick="createRule()">Create rule</button>
        </div>
        <div class="hint">Reuses the MailRuleEngine. Rules run before the AI and file at 100%. Then use <b>Rule Impact</b> to preview the effect across the corpus before it goes live. ⚠ Supplier/* rules feed the PO matchers — measure on golden first.</div>
      </div>
    </details>`;
}

async function createRule(){
  const sig = document.getElementById('rSig').value, op = document.getElementById('rOp').value;
  const val = document.getElementById('rVal').value.trim(), cat = document.getElementById('rCat').value;
  const name = document.getElementById('rName').value.trim() || (val+' → '+cat);
  if(!val){ alert('Enter a match value.'); return; }
  const by = document.getElementById('labeler').value.trim() || 'eval-ui';
  const rule = { name, enabled:true, priority:100, categoryPath:cat,
                 conditions:[{ signal:sig, operator:op, values:[val], minMatches:1 }] };
  try { await jpost('/rules?by='+encodeURIComponent(by), rule); toast('Rule created · '+name); }
  catch(e){ alert('Create failed: '+e.message); }
}

async function addCategory(){
  let path = prompt('New category path (Top/Sub, e.g. "Other/Voicemail"):', 'Other/');
  if(path === null) return;                      // cancelled
  path = path.trim().replace(/^\/+|\/+$/g,'').trim();
  if(!path){ alert('Enter a category path, e.g. "Other/Voicemail".'); return; }
  const desc = (prompt('Optional one-line description (helps the classifier target it):', '') || '').trim();
  try {
    const res = await jpost('/leaves', { categoryPath: path, description: desc });
    LEAVES = res.leaves; buildLeafPickers();
    toast(res.added ? ('Category added · '+res.path) : ('Already exists · '+res.path));
    // re-render the detail so both pickers include the new leaf, then pre-select it (Save still commits).
    if(CUR){ await pick(IDX); const sel=document.getElementById('leafPick'); if(sel) sel.value = res.path; }
  } catch(e){ alert('Add failed: '+e.message); }
}

async function save(){
  if(!CUR) return;
  const cat = document.getElementById('leafPick').value;
  const by = document.getElementById('labeler').value.trim() || 'human';
  await jpost('/golden', { mailItemId:CUR.mailItemId, goldenCategory:cat, subject:CUR.subject, fromAddress:CUR.fromAddress, labeledBy:by });
  // update local state without a full reload
  const wasBoot = isBoot(CUR);
  CUR.goldenCategory = cat; CUR.labeledBy = by;
  if(wasBoot){ GOLD.humanLabeled++; GOLD.bootstrap--; }
  document.getElementById('cLabeled').textContent = GOLD.humanLabeled;
  document.getElementById('cBoot').textContent = GOLD.bootstrap;
  const pct = GOLD.total ? Math.round(100*GOLD.humanLabeled/GOLD.total) : 0;
  document.getElementById('cPct').innerHTML = '<b>'+pct+'%</b> done';
  toast('Saved · '+cat);
  // advance to next row still in view
  const nextRows = filtered();
  if(document.getElementById('fView').value==='unlabeled'){ render(); if(nextRows.length) pick(Math.min(IDX, nextRows.length-1)); else document.getElementById('detail').innerHTML='<div class="empty">All filtered rows labeled. 🎉</div>'; }
  else { render(); pick(Math.min(IDX+1, (window._view.length-1))); }
}

document.addEventListener('keydown', e=>{ if(e.key==='Enter' && CUR && document.getElementById('leafPick')){ e.preventDefault(); save(); } });

async function runBaseline(){
  const btn = document.getElementById('runBtn');
  if(GOLD.humanLabeled === 0){ alert('No human-labeled rows yet. Correct some labels first — a run over bootstrap labels is meaningless (self-agreement).'); return; }
  if(!confirm(`Run the classifier over ${GOLD.humanLabeled} human-labeled rows? Each is one AI call.`)) return;
  btn.disabled = true; btn.textContent = 'Running…';
  try {
    await jpost('/run', { recordResponses:true });
    let snap;
    do { await new Promise(r=>setTimeout(r,1500)); snap = await jget('/status');
         btn.textContent = `Running ${snap.processed}/${snap.total}…`;
    } while(snap.running);
    btn.textContent = 'Run Baseline'; btn.disabled = false;
    openReport();
  } catch(e){ alert('Run failed: '+e.message); btn.textContent='Run Baseline'; btn.disabled=false; }
}

async function openReport(){
  const rep = await jget('/report');
  const m = document.getElementById('reportModal'); const b = document.getElementById('reportBody');
  if(!rep){ b.innerHTML = '<h2>Baseline report</h2><div class="empty">No run yet. Label some rows, then Run Baseline.</div><div style="text-align:right"><button class="ghost" onclick="closeReport()">Close</button></div>'; m.classList.add('open'); return; }
  const cls = v => v>=.9?'ok':v>=.7?'mid':'lo';
  const leafRows = rep.byLeaf.map(l=>`<tr><td>${esc(l.leaf)}</td><td class="num">${l.support}</td>
     <td class="num ${cls(l.precision)}">${(l.precision*100).toFixed(0)}%</td>
     <td class="num ${cls(l.recall)}">${(l.recall*100).toFixed(0)}%</td>
     <td class="num ${cls(l.f1)}">${(l.f1*100).toFixed(0)}%</td></tr>`).join('');
  // top confusions
  const conf = [];
  for(const g in rep.confusion) for(const p in rep.confusion[g]) if(g!==p) conf.push([g,p,rep.confusion[g][p]]);
  conf.sort((a,b)=>b[2]-a[2]);
  const confRows = conf.slice(0,12).map(c=>`<tr><td>${esc(c[0])}</td><td>${esc(c[1])}</td><td class="num">${c[2]}</td></tr>`).join('') || '<tr><td colspan="3" class="empty">No confusions 🎉</td></tr>';
  const calRows = rep.calibration.filter(c=>c.count>0).map(c=>`<tr><td>${(c.lowConfidence*100).toFixed(0)}–${(c.highConfidence*100).toFixed(0)}%</td><td class="num">${c.count}</td><td class="num">${c.accuracy!=null?(c.accuracy*100).toFixed(0)+'%':'–'}</td></tr>`).join('');
  const prov = rep.byProvider.map(p=>`${esc(p.provider)} ${p.count} (${p.count?Math.round(100*p.correct/p.count):0}%)`).join(' · ');
  b.innerHTML = `
    <h2>Baseline report</h2>
    <div class="chip" style="display:inline-block">overall <b class="${cls(rep.overallAccuracy)}">${(rep.overallAccuracy*100).toFixed(1)}%</b> · ${rep.correctItems}/${rep.totalItems} · ${esc(prov)}</div>
    <div class="sec"><h3>Per-leaf precision / recall / F1</h3>
      <table><thead><tr><th>Leaf</th><th>n</th><th>P</th><th>R</th><th>F1</th></tr></thead><tbody>${leafRows}</tbody></table></div>
    <div class="sec"><h3>Top confusions (gold → predicted)</h3>
      <table><thead><tr><th>Gold (true)</th><th>Predicted</th><th>n</th></tr></thead><tbody>${confRows}</tbody></table></div>
    <div class="sec"><h3>Confidence calibration (predicted conf → actual accuracy)</h3>
      <table><thead><tr><th>Confidence</th><th>n</th><th>Accuracy</th></tr></thead><tbody>${calRows}</tbody></table>
      <div class="hint">Where predicted-confidence stops matching actual accuracy is where the gate threshold belongs.</div></div>
    <div style="text-align:right;margin-top:14px"><button class="ghost" onclick="closeReport()">Close</button></div>`;
  m.classList.add('open');
}
function closeReport(){ document.getElementById('reportModal').classList.remove('open'); }

// ── Rule impact (deterministic re-run of the ruleset over existing items) ──────────────
function openImpact(){
  const cats = Object.keys(GOLD.byCategory||{});
  const catOpts = '<option value="">All categories</option>' + cats.map(c=>`<option>${esc(c)}</option>`).join('');
  document.getElementById('impactBody').innerHTML = `
    <h2>Rule impact <span style="color:var(--mut);font-weight:400;font-size:13px">· deterministic, no AI</span></h2>
    <div class="labelbar">
      <label>Scope <select id="impScope">${catOpts}</select></label>
      <button id="impRun" onclick="runImpact(true)">Dry-run preview</button>
      <button class="ghost" onclick="closeImpact()">Close</button>
    </div>
    <div class="hint">Re-runs the current ruleset over existing items and shows what would flip. Items no rule matches are shown unchanged (the AI isn't re-run here). Apply writes the rule result (no matcher kick); the Inbox reflects it after the next cache refresh.</div>
    <div id="impResult"></div>`;
  document.getElementById('impactModal').classList.add('open');
}
function closeImpact(){ document.getElementById('impactModal').classList.remove('open'); }

async function runImpact(dryRun){
  const scope = document.getElementById('impScope').value;
  const btn = document.getElementById('impRun'); btn.disabled=true; btn.textContent='Running…';
  try {
    await jpost('/rule-impact', { category: scope||null, dryRun });
    let s; do { await new Promise(r=>setTimeout(r,1000)); s=await jget('/rule-impact/status'); btn.textContent=`Scanning ${s.processed}/${s.total}…`; } while(s.running);
    btn.disabled=false; btn.textContent='Dry-run preview';
    renderImpact(await jget('/rule-impact/report'));
  } catch(e){ alert('Impact run failed: '+e.message); btn.disabled=false; btn.textContent='Dry-run preview'; }
}

function renderImpact(rep){
  if(!rep){ document.getElementById('impResult').innerHTML='<div class="empty">No result.</div>'; return; }
  const chRows = rep.changes.map(c=>`<tr><td>${esc(c.subject)||'(no subj)'}</td><td>${esc(c.oldCategory)}</td><td>→ ${esc(c.newCategory)}</td><td>${esc(c.rule)}</td></tr>`).join('') || '<tr><td colspan="4" class="empty">No items would change.</td></tr>';
  const distRows = Object.entries(rep.newDistribution).map(([k,v])=>`<tr><td>${esc(k)}</td><td class="num">${v}</td></tr>`).join('');
  document.getElementById('impResult').innerHTML = `
    <div class="chip" style="display:inline-block;margin-top:10px">scanned ${rep.totalScanned} · rule-matched ${rep.ruleMatched} · <b>changed ${rep.changed}</b>${rep.dryRun?' · DRY-RUN':' · applied '+rep.applied}</div>
    ${rep.changed&&rep.dryRun?`<div style="margin-top:8px"><button onclick="if(confirm('Write the rule result for ${rep.changed} items? (no matcher kick)'))runImpact(false)">Apply ${rep.changed} changes</button></div>`:''}
    <div class="sec"><h3>Would change (${rep.changes.length})</h3>
      <table><thead><tr><th>Subject</th><th>Old</th><th>New</th><th>Rule</th></tr></thead><tbody>${chRows}</tbody></table></div>
    <div class="sec"><h3>Resulting distribution</h3>
      <table><thead><tr><th>Category</th><th>n</th></tr></thead><tbody>${distRows}</tbody></table></div>`;
}

reload().catch(e=>document.getElementById('list').innerHTML='<div class="empty">Failed to load: '+esc(e.message)+'</div>');
</script>
</body>
</html>
""";
}
