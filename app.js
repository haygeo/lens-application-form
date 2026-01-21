// ===============================
// Config & Data
// ===============================

// Your Apps Script Web App URL:
const SUBMIT_ENDPOINT = "https://script.google.com/macros/s/AKfycbxRR3GogMYSWYPKtAR15TtxU1Vq19gySh4GRrzU8m71kuZPsKPnoM1ncsQucTIFv0Ub0g/exec";
const CONFIG_ENDPOINT = `${SUBMIT_ENDPOINT}?action=config`;

// Cache measures to reduce repeated prompts in preview (and for speed in prod)
const LS_MEASURES_CACHE_KEY = "rp_measures_cache";
const MEASURES_CACHE_MAX_AGE_MS = 24 * 60 * 60 * 1000; // 24h

// View state (remember where user is)
const LS_VIEW_KEY = "rp_view_state";

// Silent autosave/restore (no UI banner)
const LS_AUTOSAVE_KEY = "rp_autosave_state";
const AUTOSAVE_DEBOUNCE_MS = 600;
let autosaveTimer_ = null;

// Criteria & thresholds
const CRITERIA = [
  {id:1,name:'Crop diversity',question:'% arable land with ≥3 crops over 3 years',thresholds:{foundation:30,advanced:50,leading:70},unit:'%'},
  {id:2,name:'Soils covered',question:'% arable land covered ≥10 months/year',thresholds:{foundation:30,advanced:50,leading:70},unit:'%'},
  {id:3,name:'Cover crops',question:'% arable land with cover crops this year',thresholds:{foundation:10,advanced:30,leading:50},unit:'%'},
  {id:4,name:'Minimise soil disturbance',question:'% arable land under minimum tillage (<15 cm)',thresholds:{foundation:10,advanced:50,leading:70},unit:'%'},
  // Criteria 5 & 6: temporary yes/no mapping
  {id:5,name:'Integrated nutrient management principles',question:'Do you apply integrated nutrient management principles?',thresholds:{foundation:1,advanced:3,leading:4},unit:'yesno'},
  {id:6,name:'Integrated pest management principles',question:'Do you apply integrated pest management principles?',thresholds:{foundation:1,advanced:3,leading:5},unit:'yesno'},
  {id:7,name:'Land set aside for nature',question:'% land area set aside for nature (from nature infrastructure)',thresholds:{foundation:1,advanced:5,leading:5},unit:'%'}
];

const INFRA_LOOKUP = [
  {type:'Hedgerow', unit:'metres', impact:20},
  {type:'Row of trees', unit:'metres', impact:10},
  {type:'Isolated tree', unit:'number', impact:30},
  {type:'Woodland', unit:'m²', impact:1.5},
  {type:'Ponds', unit:'m²', impact:1.5},
  {type:'Natural ditches', unit:'metres', impact:10},
  {type:'Non-productive field margins', unit:'metres', impact:9},
  {type:'Fallow or herbal ley', unit:'m²', impact:1},
  {type:'Pollinator-friendly leys', unit:'m²', impact:1.5},
  {type:'Dry stone walls', unit:'metres', impact:1}
];

const MEASURE_GROUPS = ['in-field','capital','resilience'];
const MEASURE_CATEGORIES = [
  'Maximise plant diversity',
  'Keep soil covered / Maintain living roots in soil',
  'Minimise soil disturbance',
  'Reduce use of synthetic inputs',
  'Integrate livestock',
  'Wildlife Management',
  'Planting and managing hedegrows / woodlands',
  'Flood management',
  'Water quality and resource management',
  'Trials, training and capacity building',
  'Innovation',
  'Resilience Payment'
];

// Section 6 crop dropdown options
const CROP_OPTIONS = [
  'Winter wheat (Triticum aestivum)',
  'Winter rapeseed (Brassica napus)',
  'Spring barley (Hordeum sativum)',
  'Winter rye (Secale cereale)',
  'Winter barley (Hordeum sativum)',
  'Potatos (industry) (Solanum tuberosum)',
  'Winter oats (Avena sativa)',
  'Peas (Pisum sativum)',
  'Bean',
  'Sugar beets (Beta vulgaris)',
  'Spring wheat (Triticum aestivum)',
  'Other field crops'
];

function stablePick(list, seed){
  let h = 0;
  for(let i=0;i<seed.length;i++) h = (h*31 + seed.charCodeAt(i)) >>> 0;
  return list[h % list.length];
}

// ===============================
// State
// ===============================
// Prevent duplicate event bindings (reset() rebuilds tables)
const BOUND = { rot:false, infra:false, crops:false };

const app = {
  measuresCfg: null,
  measuresMeta: null,
  lastConfigLoadedAt: null,
  configLoadError: null,
  flow: null, // 'new' | 'trade2025'
  submitted: false,
  submitState: { new:'ready', trade2025:'ready' },
  submitDirty: { new:false, trade2025:false },
  done: new Set(),
  warn: new Set(),
  data: {
    applicant: { name:'', business:'', email:'', type:'' },
    baseline: {
      agricultural_ha:null,
      arable_ha:null,
      perm_cropland_ha:null,
      perm_pasture_ha:null,
      habitat_ha:null,
      livestock:null,
      fuel_lpy:null
    },
    rotations: [],
    infrastructure: [],
    criteria_inputs: { c2:null, c3:null, c4:null, c5:false, c6:false },
    crops: [],
    mrvCriteria: {} // C1..C7 for existing farms
  }
};

const SUBMIT_LABELS = {
  ready: 'Submit and show eligible measures',
  loading: 'Submitting...',
  submitted: 'Submitted ✓',
  resubmit: 'Resubmit',
  error: 'Try again'
};

function submitKey_(flow){
  return flow === 'new' ? 'new' : 'trade2025';
}

function setSubmitState_(flow, state){
  const key = submitKey_(flow);
  app.submitState[key] = state;
  const btnId = flow === 'new' ? 'submit-new' : 'submit-existing';
  const btn = document.getElementById(btnId);
  if(!btn) return;
  const label = btn.querySelector('.btn-label');
  if(label) label.textContent = SUBMIT_LABELS[state] || SUBMIT_LABELS.ready;
  btn.classList.remove('status-ready','status-loading','status-submitted','status-resubmit','status-error');
  btn.classList.add(`status-${state}`);
}

function syncSubmitState_(flow){
  const key = submitKey_(flow);
  if(app.submitState[key] === 'loading') return;
  const state = app.submitted ? (app.submitDirty[key] ? 'resubmit' : 'submitted') : 'ready';
  setSubmitState_(flow, state);
}

function markSubmitDirty_(flow){
  const key = submitKey_(flow);
  if(app.flow !== flow || !app.submitted || app.submitDirty[key]) return;
  app.submitDirty[key] = true;
  setSubmitState_(flow, 'resubmit');
}

function resetSubmitState_(){
  app.submitState = { new:'ready', trade2025:'ready' };
  app.submitDirty = { new:false, trade2025:false };
  setSubmitState_('new', 'ready');
  setSubmitState_('trade2025', 'ready');
}

// ===============================
// Missing highlight helpers
// ===============================
function setMissing(el, on){
  if(!el) return;
  el.classList.toggle('missing', !!on);
}

function clearMissingWithin(root){
  if(!root) return;
  root.querySelectorAll('.missing').forEach(x=> x.classList.remove('missing'));
  root.querySelectorAll('.missing-block').forEach(x=> x.classList.remove('missing-block'));
}

// ===============================
// Measures config (fetch first, fallback JSONP)
// ===============================
function setModalStatus_(msg){
  const el = document.getElementById('measures-modal-status');
  if(el) el.textContent = msg || '';
}

function writeMeasuresCache_(measures){
  try{ localStorage.setItem(LS_MEASURES_CACHE_KEY, JSON.stringify({ ts: Date.now(), measures })); } catch {}
}

function readMeasuresCache_(){
  try{
    const raw = localStorage.getItem(LS_MEASURES_CACHE_KEY);
    if(!raw) return null;
    const obj = JSON.parse(raw);
    if(!obj || !obj.ts || !Array.isArray(obj.measures)) return null;
    if(Date.now() - obj.ts > MEASURES_CACHE_MAX_AGE_MS) return null;
    return obj;
  } catch { return null; }
}

function isBoolTrue(v){
  if(v === true) return true;
  if(v === false) return false;
  const s = String(v ?? '').trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === '1';
}

function normalizeGroup_(g){
  const s = String(g ?? '').toLowerCase().trim();
  if(!s) return 'in-field';
  if(s === 'infield' || s === 'in field' || s === 'in-field' || s === 'in_field') return 'in-field';
  if(s === 'capital') return 'capital';
  if(s === 'resilience' || s === 'resilience payment') return 'resilience';
  return s;
}

function cfgFromSheetMeasures(sheetMeasures){
  const cfg = {};
  const meta = {};

  for(const m of sheetMeasures){
    const code = (m.measure_code || m.code || '').trim();
    const name = (m.measure_name || m.name || '').trim();
    if(!code) continue;

    meta[code] = { code, name };

    const criteria = {};
    for(let i=1;i<=7;i++) criteria[i] = isBoolTrue(m['C'+i]);

    const rawGroup = (m.measure_group || m.group || stablePick(MEASURE_GROUPS, code));
    const rawCat = (m.measure_category || m.category || stablePick(MEASURE_CATEGORIES, code+'cat'));

    cfg[code] = {
      active: isBoolTrue(m.active),
      group: normalizeGroup_(rawGroup),
      category: String(rawCat ?? '').trim(),
      criteria
    };
  }

  return { cfg, meta };
}

function applyMeasuresFromSheet_(sheetMeasures, sourceLabel){
  const { cfg, meta } = cfgFromSheetMeasures(sheetMeasures);
  if(Object.keys(cfg).length === 0) throw new Error('No measures found in the Google Sheet Measures tab.');

  app.measuresCfg = cfg;
  app.measuresMeta = meta;
  app.lastConfigLoadedAt = new Date();
  app.configLoadError = null;
  setModalStatus_(`${sourceLabel}: loaded ${Object.keys(cfg).length} measures.`);
}

function loadJsonp(url, callbackName, timeoutMs=12000){
  return new Promise((resolve, reject)=>{
    const cb = callbackName || ('__jsonp_cb_' + Math.random().toString(36).slice(2));
    const script = document.createElement('script');
    const timer = setTimeout(()=>{ cleanup(); reject(new Error('JSONP request timed out')); }, timeoutMs);

    function cleanup(){
      clearTimeout(timer);
      if(script && script.parentNode) script.parentNode.removeChild(script);
      try{ delete window[cb]; } catch { window[cb] = undefined; }
    }

    window[cb] = (data)=>{ cleanup(); resolve(data); };

    const sep = url.includes('?') ? '&' : '?';
    script.src = `${url}${sep}callback=${encodeURIComponent(cb)}&_ts=${Date.now()}`;
    script.onerror = ()=>{ cleanup(); reject(new Error('JSONP script failed to load')); };
    document.head.appendChild(script);
  });
}

async function loadConfigFromSheet(force=false){
  setModalStatus_('');

  if(!force){
    const cached = readMeasuresCache_();
    if(cached){
      try{ applyMeasuresFromSheet_(cached.measures, 'Using cached measures'); return; } catch {}
    }
  }

  try{
    setModalStatus_('Connecting to Google Sheet…');
    const controller = new AbortController();
    const timeout = setTimeout(()=>controller.abort(), 15000);
    const res = await fetch(CONFIG_ENDPOINT + `&_ts=${Date.now()}`, { method:'GET', cache:'no-store', signal: controller.signal });
    clearTimeout(timeout);
    if(!res.ok) throw new Error(`Config fetch failed (HTTP ${res.status})`);
    const json = await res.json();
    if(!json || json.status !== 'ok' || !Array.isArray(json.measures)) throw new Error('Config endpoint did not return expected JSON.');

    applyMeasuresFromSheet_(json.measures, 'Loaded from Google Sheet (fetch)');
    writeMeasuresCache_(json.measures);

  } catch (e) {
    console.warn('Fetch config failed; trying JSONP fallback:', e);
    try{
      setModalStatus_('Fetch blocked — trying fallback…');
      const jsonp = await loadJsonp(CONFIG_ENDPOINT, null, 15000);
      if(!jsonp || jsonp.status !== 'ok' || !Array.isArray(jsonp.measures)) throw new Error('JSONP fallback did not return expected data.');

      applyMeasuresFromSheet_(jsonp.measures, 'Loaded from Google Sheet (fallback)');
      writeMeasuresCache_(jsonp.measures);

    } catch (e2) {
      console.warn('Config load failed:', e2);
      app.configLoadError = String(e2);
      setModalStatus_(`⚠ Could not load measures from Google Sheet: ${String(e2)}`);
      app.measuresCfg = null;
      app.measuresMeta = null;
    }
  }
}

function allActiveMeasuresForLists_(){
  if(!app.measuresCfg || !app.measuresMeta) return [];
  const out = [];
  for(const code of Object.keys(app.measuresCfg)){
    const c = app.measuresCfg[code];
    if(!c || !c.active) continue;
    out.push({
      code,
      name: app.measuresMeta[code]?.name || '',
      group: c.group,
      category: c.category,
      contributes_to: Array.from({length:7},(_,i)=>i+1).filter(i=> !!(c.criteria && c.criteria[i])),
      criteriaFlags: c.criteria || {}
    });
  }
  out.sort((a,b)=> a.code.localeCompare(b.code));
  return out;
}

function effectiveMeasures(){
  // For eligibility calculations, we only include measures that contribute to at least one criterion.
  return allActiveMeasuresForLists_().filter(m => (m.contributes_to || []).length > 0);
}

// ===============================
// Pathway logic helpers
// ===============================
function rangeText(t, unit){
  if(unit === 'yesno'){
    return `
      <div class="ranges">
        <span class="range entry"><strong>No</strong>: Below-entry</span>
        <span class="range engaged"><strong>Yes</strong>: Entry (placeholder)</span>
      </div>
    `;
  }

  const f=t.foundation??0, a=t.advanced??Infinity, l=t.leading??Infinity;
  const u=unit==='%'?'%':'';

  const engagedEnd = (a===Infinity) ? null : (a-1);
  const advancedEnd = (l===Infinity) ? null : (l-1);

  const engagedTxt = engagedEnd == null ? `${f}${u}–∞` : `${f}${u}–${engagedEnd}${u}`;

  let advancedTxt = '—';
  if(a !== Infinity && l !== Infinity && advancedEnd >= a){
    advancedTxt = `${a}${u}–${advancedEnd}${u}`;
  } else if(a !== Infinity && l === Infinity){
    advancedTxt = `${a}${u}–∞`;
  }

  const leadingTxt = (l===Infinity) ? '—' : `≥ ${l}${u}`;

  return `
    <div class="ranges">
      <span class="range entry"><strong>Entry</strong>: &lt; ${f}${u}</span>
      <span class="range engaged"><strong>Engaged</strong>: ${engagedTxt}</span>
      <span class="range advanced"><strong>Advanced</strong>: ${advancedTxt}</span>
      <span class="range leading"><strong>Leading</strong>: ${leadingTxt}</span>
    </div>
  `;
}

function levelFor(value, t){
  const f=t.foundation??0, a=t.advanced??Infinity, l=t.leading??Infinity;
  if(value==null||isNaN(value)) return null;
  if(value<f) return 'entry';
  if(value<a) return 'engaged';
  if(value<l) return 'advanced';
  return 'leading';
}

function minLevel(levels){
  const order=['entry','engaged','advanced','leading'];
  return levels.reduce((m,lv)=> order.indexOf(lv)<order.indexOf(m)?lv:m,'leading');
}

function suggestedMeasures(measures, levelsByCrit){
  if(!measures || measures.length === 0) return [];
  const weak=new Set(Object.entries(levelsByCrit)
    .filter(([,v])=>v==='entry'||v==='engaged')
    .map(([k])=>+k));
  return measures.filter(m=>m.contributes_to.some(c=>weak.has(c)));
}

function groupMeasures(measures){
  const out = new Map();
  for(const m of measures){
    const g = normalizeGroup_(m.group || 'in-field');
    const cat = m.category || 'Uncategorised';
    if(!out.has(g)) out.set(g, new Map());
    const catMap = out.get(g);
    if(!catMap.has(cat)) catMap.set(cat, []);
    catMap.get(cat).push(m);
  }
  for(const [,catMap] of out){
    for(const [cat,list] of catMap){
      list.sort((a,b)=> a.code.localeCompare(b.code));
      catMap.set(cat, list);
    }
  }
  return out;
}

function renderMeasure(m){
  return `<div class="measure">
    <div class="flex" style="justify-content:space-between;align-items:flex-start;gap:10px">
      <div class="grow">
        <h4>${escapeHtml(m.name || '')}</h4>
        <div class="small">Code: <span class="mono">${escapeHtml(m.code)}</span></div>
      </div>
    </div>
  </div>`;
}

function renderMeasureIneligible(m){
  return `<div class="measure ineligible">
    <div class="flex" style="justify-content:space-between;align-items:flex-start;gap:10px">
      <div class="grow">
        <h4>${escapeHtml(m.name || '')}</h4>
        <div class="small">Code: <span class="mono">${escapeHtml(m.code)}</span></div>
      </div>
      <span class="chip-ineligible">Ineligible</span>
    </div>
  </div>`;
}

function renderMeasuresThreeColumns(measuresEl, measures){
  measuresEl.innerHTML='';
  if(!measures || measures.length===0){
    measuresEl.innerHTML = `<div class="small">No measures to display yet.</div>`;
    return;
  }

  const grouped = groupMeasures(measures);
  const groupOrder = ['in-field','capital','resilience'];

  const cols = document.createElement('div');
  cols.className = 'measures-columns';

  for(const g of groupOrder){
    const col = document.createElement('div');
    col.className = 'measure-col';
    col.dataset.group = g;

    const title = (g === 'in-field') ? 'In-field' : (g[0].toUpperCase()+g.slice(1));
    col.innerHTML = `<div class="col-head"><div class="col-title">${escapeHtml(title)}</div><div class="small">Eligible measures</div></div>`;

    const catMap = grouped.get(g);
    if(!catMap || catMap.size === 0){
      col.insertAdjacentHTML('beforeend', `<div class="small">No eligible ${escapeHtml(title)} measures.</div>`);
      cols.appendChild(col);
      continue;
    }

    const cats = [...catMap.keys()].sort((a,b)=>{
      if(a==='Resilience Payment') return 1;
      if(b==='Resilience Payment') return -1;
      return a.localeCompare(b);
    });

    for(const cat of cats){
      const block = document.createElement('div');
      block.className = 'cat-block';
      block.innerHTML = `<div class="cat-title">${escapeHtml(cat)}</div>`;
      for(const m of catMap.get(cat)) block.insertAdjacentHTML('beforeend', renderMeasure(m));
      col.appendChild(block);
    }

    cols.appendChild(col);
  }

  measuresEl.appendChild(cols);
}

function renderMeasuresThreeColumnsIneligible(measuresEl, measures){
  measuresEl.innerHTML='';
  if(!measures || measures.length===0){
    measuresEl.innerHTML = `<div class="small">No ineligible measures to display.</div>`;
    return;
  }

  const grouped = groupMeasures(measures);
  const groupOrder = ['in-field','capital','resilience'];

  const cols = document.createElement('div');
  cols.className = 'measures-columns';

  for(const g of groupOrder){
    const col = document.createElement('div');
    col.className = 'measure-col ineligible';
    col.dataset.group = g;

    const title = (g === 'in-field') ? 'In-field' : (g[0].toUpperCase()+g.slice(1));
    col.innerHTML = `<div class="col-head"><div class="col-title">${escapeHtml(title)}</div><div class="small">Ineligible measures</div></div>`;

    const catMap = grouped.get(g);
    if(!catMap || catMap.size === 0){
      col.insertAdjacentHTML('beforeend', `<div class="small">No ineligible ${escapeHtml(title)} measures.</div>`);
      cols.appendChild(col);
      continue;
    }

    const cats = [...catMap.keys()].sort((a,b)=> a.localeCompare(b));

    for(const cat of cats){
      const block = document.createElement('div');
      block.className = 'cat-block';
      block.innerHTML = `<div class="cat-title">${escapeHtml(cat)}</div>`;
      for(const m of catMap.get(cat)) block.insertAdjacentHTML('beforeend', renderMeasureIneligible(m));
      col.appendChild(block);
    }

    cols.appendChild(col);
  }

  measuresEl.appendChild(cols);
}

function renderResiliencePayment(){
  return {
    code:'RES_PAY',
    name:'Resilience Payment',
    group:'resilience',
    category:'Resilience Payment',
    contributes_to:[1,2,3,4,5,6,7],
    criteriaFlags: {1:true,2:true,3:true,4:true,5:true,6:true,7:true}
  };
}

function escapeHtml(str){
  return String(str ?? '')
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#039;');
}

function setAriaDisabled_(el, disabled){
  if(!el) return;
  el.setAttribute('aria-disabled', disabled ? 'true' : 'false');
  if('disabled' in el) el.disabled = !!disabled;
  el.classList.toggle('disabled', !!disabled);
}

// ===============================
// Progress tracker
// ===============================
function markDone(stepKey){ app.done.add(stepKey); }

function setStepper(active){
  const steps = [
    {key:'s1', label:'1. Identification'},
    ...(app.flow === 'new' ? [
      {key:'s2', label:'2. Baseline'},
      {key:'s3', label:'3. Rotations'},
      {key:'s4', label:'4. Nature infra'},
      {key:'s5', label:'5. Criteria'},
      {key:'s6', label:'6. Crops'},
      {key:'s7', label:'7. Review'}
    ] : []),
    ...(app.flow === 'trade2025' ? [{key:'mrv', label:'MRV criteria'}] : []),
  ];

  const el = document.getElementById('stepper');
  if(!el) return;

  el.innerHTML = steps.map((s,i)=>{
    const isActive = s.key === active;
    const isDone = app.done.has(s.key);
    const isWarn = app.warn.has(s.key);
    return `<div class="step ${isDone?'done':''} ${isWarn?'warn':''} ${isActive?'active':''}"><span class="num">${i+1}</span>${escapeHtml(s.label)}</div>`;
  }).join('');
}

function updateStepperStatusesNew_(showMissing){
  const baselineReq = ['s2-agri','s2-arable','s2-perm-crop','s2-pasture','s2-habitat','s2-livestock','s2-fuel'];
  const baselineMissing = baselineReq.filter(id => {
    const el = document.getElementById(id);
    return !el || String(el.value||'').trim()==='';
  });

  if(showMissing){
    baselineReq.forEach(id=> setMissing(document.getElementById(id), baselineMissing.includes(id)));
  }

  const arable = app.data.baseline.arable_ha;
  const rotTotal = app.data.rotations.reduce((s,r)=> s + (r.area_ha || 0), 0);
  const rotHasOne = app.data.rotations.length > 0;
  const rotTooMuch = (arable!=null && arable>0 && rotTotal > arable + 1e-9);
  if(showMissing && (!rotHasOne || rotTooMuch)) document.getElementById('rot-block')?.classList.add('missing-block');

  const meta = infraMeta();
  const infraOk = meta.agriM2 != null;
  if(showMissing && !infraOk) document.getElementById('infra-block')?.classList.add('missing-block');

  const c = computeCriteriaForNew();
  const requiredCrit = [1,2,3,4,7];
  const critMissing = requiredCrit.filter(k => c[k] == null);
  if(showMissing) validateCriteriaSection(true);

  const cropsOk = app.data.crops.length > 0;
  if(showMissing && !cropsOk) document.getElementById('crops-block')?.classList.add('missing-block');

  const statuses = {
    s1: validateSection1(false).ok,
    s2: baselineMissing.length === 0,
    s3: rotHasOne && !rotTooMuch,
    s4: infraOk,
    s5: critMissing.length === 0,
    s6: cropsOk,
    s7: true
  };

  // Set done/warn
  app.warn = new Set();
  for(const k of Object.keys(statuses)){
    if(statuses[k]) app.done.add(k);
    else { app.done.delete(k); app.warn.add(k); }
  }

  setStepper('s7');
  return statuses;
}

function updateStepperStatusesExisting_(showMissing){
  const s1ok = validateSection1(false).ok;
  if(showMissing) validateExisting(true);

  const required = [1,2,3,4,7];
  const mrvMissing = required.filter(i => app.data.mrvCriteria['C'+i] == null);
  const mrvOk = mrvMissing.length === 0;

  app.warn = new Set();
  if(s1ok) app.done.add('s1'); else { app.done.delete('s1'); app.warn.add('s1'); }
  if(mrvOk) app.done.add('mrv'); else { app.done.delete('mrv'); app.warn.add('mrv'); }

  setStepper('mrv');
  return { s1: s1ok, mrv: mrvOk };
}

function markAllDoneForFlow_(){
  app.warn = new Set();
  if(app.flow === 'new') ['s1','s2','s3','s4','s5','s6','s7'].forEach(k=> app.done.add(k));
  if(app.flow === 'trade2025') ['s1','mrv'].forEach(k=> app.done.add(k));
}

// ===============================
// Full Measures List Modal
// ===============================
function renderCriteriaLegend_(){
  // We use a simple table now, so we don't need the pill legend.
  const el = document.getElementById('criteria-legend');
  if(el) el.innerHTML = '';
}

function renderCriteriaGridForMeasure_(criteriaFlags){
  const cells = [];
  for(let i=1;i<=7;i++){
    const ok = !!(criteriaFlags && criteriaFlags[i]);
    cells.push(`<div class="crit-cell">${ok ? '✓' : '<span class="x">✕</span>'}</div>`);
  }
  return `<div class="crit-grid">${cells.join('')}</div>`;
}

function renderMeasureForModal_(m){
  return `<div class="measure">
    <div class="flex" style="justify-content:space-between;align-items:flex-start;gap:10px">
      <div class="grow">
        <h4>${escapeHtml(m.name || '')}</h4>
        <div class="small">Code: <span class="mono">${escapeHtml(m.code)}</span></div>
        <div class="small" style="margin-top:6px"><strong>Criteria contribution</strong> (✓ contributes, ✕ does not)</div>
        ${renderCriteriaGridForMeasure_(m.criteriaFlags)}
      </div>
    </div>
  </div>`;
}

function renderMeasuresModal_(){
  const body = document.getElementById('measures-modal-body');
  if(!body) return;

  if(!app.measuresCfg || !app.measuresMeta){
    body.innerHTML = `<div class="error"><strong>Measures could not be loaded.</strong><div class="small" style="margin-top:6px">Please check your Google Apps Script access is set to <strong>Anyone</strong>, then refresh this page.</div></div>`;
    return;
  }

  const measures = allActiveMeasuresForLists_(); // already active + sorted by code

  const header = `
    <div class="measures-table-wrap">
      <table class="measures-table">
        <thead>
          <tr>
            <th>Measure code</th>
            <th>Measure name</th>
            <th>Group</th>
            <th>Category</th>
            <th>C1</th><th>C2</th><th>C3</th><th>C4</th><th>C5</th><th>C6</th><th>C7</th>
          </tr>
        </thead>
        <tbody>
  `;

  const rows = measures.map(m=>{
    const g = normalizeGroup_(m.group);
    const groupLabel = (g === 'in-field') ? 'In-field' : (g[0].toUpperCase()+g.slice(1));
    const flags = m.criteriaFlags || {};
    const cell = (i)=> flags[i] ? `<span class="tick">✓</span>` : `<span class="cross">✕</span>`;
    return `
      <tr>
        <td class="mono">${escapeHtml(m.code)}</td>
        <td class="wrap">${escapeHtml(m.name || '')}</td>
        <td>${escapeHtml(groupLabel)}</td>
        <td class="wrap">${escapeHtml(m.category || '')}</td>
        <td>${cell(1)}</td><td>${cell(2)}</td><td>${cell(3)}</td><td>${cell(4)}</td><td>${cell(5)}</td><td>${cell(6)}</td><td>${cell(7)}</td>
      </tr>
    `;
  }).join('');

  const footer = `
        </tbody>
      </table>
    </div>
  `;

  body.innerHTML = header + rows + footer;
}

function openMeasuresModal_(){
  const modal = document.getElementById('measures-modal');
  if(!modal) return;
  app._lastFocusEl = (document.activeElement instanceof HTMLElement) ? document.activeElement : null;
  document.body.classList.add('modal-open');
  modal.classList.remove('hide');
  renderCriteriaLegend_();
  renderMeasuresModal_();
  document.getElementById('measures-modal-close')?.focus();
}

function closeMeasuresModal_(){
  const modal = document.getElementById('measures-modal');
  if(!modal) return;
  if(modal.classList.contains('hide')) return;
  modal.classList.add('hide');
  document.body.classList.remove('modal-open');
  app._lastFocusEl?.focus?.();
}

function bindMeasuresModal_(){
  document.getElementById('open-measures-btn')?.addEventListener('click', async ()=>{
    // Ensure measures are loaded (or at least attempted) before opening
    if(!app.measuresCfg || !app.measuresMeta){
      await loadConfigFromSheet(true);
    }
    openMeasuresModal_();
  });

  document.getElementById('measures-modal-close')?.addEventListener('click', closeMeasuresModal_);

  // click outside closes
  document.getElementById('measures-modal')?.addEventListener('click', (e)=>{
    if(e.target?.id === 'measures-modal') closeMeasuresModal_();
  });

  // Esc closes + tab trap
  document.addEventListener('keydown', (e)=>{
    const modal = document.getElementById('measures-modal');
    if(!modal || modal.classList.contains('hide')) return;

    if(e.key === 'Escape'){
      closeMeasuresModal_();
      return;
    }
    if(e.key !== 'Tab') return;

    const focusables = modal.querySelectorAll(
      'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
    );
    if(!focusables.length) return;

    const first = focusables[0];
    const last = focusables[focusables.length - 1];
    const active = document.activeElement;

    if(e.shiftKey && active === first){
      e.preventDefault();
      last.focus();
    } else if(!e.shiftKey && active === last){
      e.preventDefault();
      first.focus();
    }
  });
}

async function ensureMeasuresLoaded_(){
  if(app.measuresCfg && app.measuresMeta) return true;
  await loadConfigFromSheet(true);
  return !!(app.measuresCfg && app.measuresMeta);
}

// ===============================
// UI rendering (criteria cards)
// ===============================
function renderCriteriaCards(container, mode){
  container.innerHTML='';
  for(const c of CRITERIA){
    const isYesNo = c.unit === 'yesno';
    const isCalculated = (mode==='new' && (c.id===1 || c.id===7));

    const disabledAttr = (mode==='new' && isCalculated) ? 'disabled' : '';

    const inputHtml = isYesNo
      ? `<label class="small" style="display:block;margin-top:10px">
           <input type="checkbox" id="crit-${mode}-${c.id}" ${disabledAttr}>
           Yes
         </label>`
      : `<div class="input-row" style="margin-top:10px">
           <input id="crit-${mode}-${c.id}" type="number" ${disabledAttr} placeholder="${isCalculated?'Calculated':''}" min="0" step="1">
           <span>${c.unit==='%'?'%':''}</span>
           <span class="badge" id="badge-${mode}-${c.id}"><span class="dot"></span><span class="level" id="badge-text-${mode}-${c.id}">—</span></span>
         </div>`;

    container.insertAdjacentHTML('beforeend',`
      <div class="card">
        <div class="section-title">Criterion C${c.id}${isCalculated ? ' (calculated)' : ''}</div>
        <h3>${escapeHtml(c.name)}</h3>
        <div class="small">${escapeHtml(c.question)}</div>
        ${rangeText(c.thresholds,c.unit)}
        ${inputHtml}
        ${isYesNo ? `<div class="input-row" style="margin-top:10px"><span class="badge" id="badge-${mode}-${c.id}"><span class="dot"></span><span class="level" id="badge-text-${mode}-${c.id}">—</span></span></div>` : ''}
        ${isCalculated && mode==='new' ? `<div class="small" style="margin-top:8px">This is calculated from earlier sections.</div>` : ''}
      </div>
    `);
  }
}

function setBadge(mode, cid, lv){
  const badge = document.getElementById(`badge-${mode}-${cid}`);
  const text = document.getElementById(`badge-text-${mode}-${cid}`);
  if(badge && text){
    badge.className = `badge ${lv||''}`;
    text.textContent = lv ? (lv[0].toUpperCase()+lv.slice(1)) : '—';
  }
}

// ===============================
// Section calculations
// ===============================
function numVal(id){
  const el = document.getElementById(id);
  if(!el) return null;
  const raw = String(el.value || '').trim();
  if(raw === '') return null;
  const v = Number(raw);
  return Number.isFinite(v) ? v : null;
}

function clampPct(x){
  if(x == null || !Number.isFinite(x)) return null;
  if(x < 0) return 0;
  if(x > 100) return 100;
  return Math.round(x * 100) / 100;
}

function calcRotationCriterion1(){
  const arable = app.data.baseline.arable_ha;
  if(!arable || arable <= 0) return null;

  let sum = 0;
  for(const r of app.data.rotations){
    if((r.num_crops || 0) >= 3) sum += (r.area_ha || 0);
  }
  const pct = (sum / arable) * 100;
  return clampPct(pct);
}

function infraMeta(){
  const agri = app.data.baseline.agricultural_ha;
  if(!agri || agri <= 0) return { totalImpactM2: 0, agriM2: null, pct: null };

  let total = 0;
  for(const row of app.data.infrastructure){
    total += (row.impact_m2 || 0);
  }
  const agriM2 = agri * 10000;
  const pct = (total / agriM2) * 100;
  return { totalImpactM2: total, agriM2, pct: clampPct(pct) };
}

function computeCriteriaForNew(){
  const c1 = calcRotationCriterion1();
  const c7 = infraMeta().pct;
  const c2 = app.data.criteria_inputs.c2;
  const c3 = app.data.criteria_inputs.c3;
  const c4 = app.data.criteria_inputs.c4;
  const c5 = app.data.criteria_inputs.c5 ? 1 : 0;
  const c6 = app.data.criteria_inputs.c6 ? 1 : 0;
  return { 1:c1, 2:c2, 3:c3, 4:c4, 5:c5, 6:c6, 7:c7 };
}

function computeCriteriaForExisting(){
  const out = {};
  for(let i=1;i<=7;i++){
    const v = app.data.mrvCriteria['C'+i];
    out[i] = v == null ? null : v;
  }
  return out;
}

function criteriaLevelsFromValues(values){
  const levels = {};
  for(const c of CRITERIA){
    const v = values[c.id];
    const lv = (v==null) ? null : levelFor(v, c.thresholds);
    levels[c.id] = lv;
  }
  return levels;
}

function overallFromLevels(levels){
  const nonNull = Object.values(levels).filter(Boolean);
  return nonNull.length ? minLevel(nonNull) : null;
}

// ===============================
// Measures selection
// ===============================
function selectEligibleMeasures(criteriaValues){
  const levels = criteriaLevelsFromValues(criteriaValues);
  const overall = overallFromLevels(levels);

  // Eligible measures
  let eligible = [];
  if(overall === 'advanced' || overall === 'leading'){
    eligible = [renderResiliencePayment()];
  } else {
    const all = effectiveMeasures();
    eligible = suggestedMeasures(all, levels);
  }

  // Ineligible measures = all active measures not in eligible
  const eligibleCodes = new Set(eligible.map(m=>m.code));
  const allActive = effectiveMeasures();
  const ineligible = allActive.filter(m=> !eligibleCodes.has(m.code));

  return { overall, levels, measures: eligible, ineligible };
}

function criteriaSummaryTable(values, levels){
  const badgeHtml = (lv)=>{
    if(!lv) return '—';
    const label = lv[0].toUpperCase() + lv.slice(1);
    return `<span class="badge ${lv}"><span class="dot"></span><span class="level">${label}</span></span>`;
  };

  const rows = CRITERIA.map(c=>{
    const v = values[c.id];
    const lv = levels[c.id];
    const disp = (v==null)
      ? '—'
      : (c.unit === 'yesno' ? (v>=1 ? 'Yes' : 'No') : `${v}${c.unit==='%'?'%':''}`);

    return `<tr>
      <td class="mono">C${c.id}</td>
      <td>${escapeHtml(c.name)}</td>
      <td>${escapeHtml(disp)}</td>
      <td>${badgeHtml(lv)}</td>
    </tr>`;
  }).join('');

  return `<div style="overflow:auto"><table class="table">
    <thead><tr><th>Criterion</th><th>What it measures</th><th>Your value</th><th>Level</th></tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;
}

// ===============================
// Tables (repeatable rows)
// ===============================
function rotRowTemplate(rowId){
  return `<tr data-id="${rowId}">
    <td><input type="text" placeholder="e.g. 4-year rotation" data-k="name"></td>
    <td><input type="number" min="0" step="1" placeholder="e.g. 4" data-k="num_crops"></td>
    <td><div class="input-row" style="margin-top:0"><input type="number" min="0" step="0.01" placeholder="e.g. 20" data-k="area_ha"><span>ha</span></div></td>
    <td><a class="btn ghost" data-act="del">Remove</a></td>
  </tr>`;
}

function infraRowTemplate(rowId){
  const options = INFRA_LOOKUP.map(x=>`<option value="${escapeHtml(x.type)}">${escapeHtml(x.type)}</option>`).join('');
  return `<tr data-id="${rowId}">
    <td><select data-k="type"><option value="" selected disabled>Select…</option>${options}</select></td>
    <td><input type="number" min="0" step="0.01" placeholder="e.g. 100" data-k="qty"></td>
    <td class="mono" data-k="unit">—</td>
    <td class="mono" data-k="impact">—</td>
    <td class="mono" data-k="impact_m2">—</td>
    <td><a class="btn ghost" data-act="del">Remove</a></td>
  </tr>`;
}

function cropRowTemplate(rowId){
  const options = CROP_OPTIONS.map(c=>`<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  return `<tr data-id="${rowId}">
    <td>
      <select data-k="crop">
        <option value="" selected disabled>Select…</option>
        ${options}
      </select>
    </td>
    <td><input type="number" min="0" step="0.01" placeholder="e.g. 10" data-k="area_ha"></td>
    <td><input type="number" min="0" step="0.01" placeholder="e.g. 120" data-k="n_kg_ha"></td>
    <td><input type="number" min="0" step="0.01" placeholder="e.g. 20" data-k="org_n_kg_ha"></td>
    <td><input type="number" min="0" step="0.01" placeholder="e.g. 8" data-k="yield_t_ha"></td>
    <td><a class="btn ghost" data-act="del">Remove</a></td>
  </tr>`;
}

function uid(prefix){ return `${prefix}_${Math.random().toString(36).slice(2,9)}`; }

function bindRotations(){
  const tbody = document.querySelector('#rot-table tbody');
  const addBtn = document.getElementById('rot-add');
  const summary = document.getElementById('rot-summary');
  const status = document.getElementById('s3-status');
  const rotBlock = document.getElementById('rot-block');

  function syncFromDom(){
    const rows = [...tbody.querySelectorAll('tr')];
    const data = [];
    for(const tr of rows){
      const name = tr.querySelector('[data-k="name"]').value.trim();
      const numCrops = Number(tr.querySelector('[data-k="num_crops"]').value || 0);
      const areaHa = Number(tr.querySelector('[data-k="area_ha"]').value || 0);
      data.push({ id: tr.dataset.id, name, num_crops: Number.isFinite(numCrops)?numCrops:null, area_ha: Number.isFinite(areaHa)?areaHa:null });
    }
    app.data.rotations = data;
  }

  function recalc(){
    syncFromDom();

    const arable = app.data.baseline.arable_ha;
    const total = app.data.rotations.reduce((s,r)=> s + (r.area_ha || 0), 0);

    if(summary){
      summary.textContent = `Total rotation area: ${Math.round(total*100)/100} ha${arable!=null ? ` • Arable: ${arable} ha` : ''}`;
    }

    const tooMuch = (arable!=null && arable>0 && total > arable + 1e-9);
    if(status){
      if(tooMuch){
        status.textContent = 'Total rotation surface area must not exceed arable cropland area.';
        status.classList.remove('warn');
        status.classList.add('err');
      } else {
        status.textContent = '';
        status.classList.remove('err');
      }
    }

    const c1 = calcRotationCriterion1();
    setCriterionValue('new', 1, c1);

    const next = document.getElementById('s3-next');
    const hasOne = app.data.rotations.length > 0;
    const ok = hasOne && !tooMuch;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    if(rotBlock){
      rotBlock.classList.toggle('missing-block', !hasOne);
    }

    validateCriteriaSection(false);
    updateReview();
  }

  if(tbody && tbody.children.length === 0) tbody.insertAdjacentHTML('beforeend', rotRowTemplate(uid('rot')));

  if(BOUND.rot){
    recalc();
    return;
  }
  BOUND.rot = true;

  addBtn.onclick = ()=>{
    markSubmitDirty_('new');
    tbody.insertAdjacentHTML('beforeend', rotRowTemplate(uid('rot')));
    recalc();
  };

  tbody.addEventListener('click', (e)=>{
    const btn = e.target.closest('[data-act="del"]');
    if(!btn) return;
    markSubmitDirty_('new');
    btn.closest('tr')?.remove();
    recalc();
  });

  tbody.addEventListener('input', ()=>{
    markSubmitDirty_('new');
    recalc();
  });

  recalc();
}

function bindInfrastructure(){
  const tbody = document.querySelector('#infra-table tbody');
  const addBtn = document.getElementById('infra-add');
  const summary = document.getElementById('infra-summary');
  const status = document.getElementById('s4-status');

  function lookup(type){
    return INFRA_LOOKUP.find(x=> x.type === type) || null;
  }

  function syncFromDom(){
    const rows = [...tbody.querySelectorAll('tr')];
    const data = [];
    for(const tr of rows){
      const type = tr.querySelector('[data-k="type"]').value;
      const qty = Number(tr.querySelector('[data-k="qty"]').value || 0);
      const info = lookup(type);
      const unit = info ? info.unit : null;
      const impact = info ? info.impact : null;
      const impactM2 = (info && Number.isFinite(qty)) ? qty * info.impact : null;

      data.push({ id: tr.dataset.id, type, qty: Number.isFinite(qty)?qty:null, unit, impact, impact_m2: impactM2 });

      tr.querySelector('[data-k="unit"]').textContent = unit ?? '—';
      tr.querySelector('[data-k="impact"]').textContent = impact ?? '—';
      tr.querySelector('[data-k="impact_m2"]').textContent = (impactM2==null ? '—' : Math.round(impactM2*100)/100);
    }
    app.data.infrastructure = data;
  }

  function recalc(){
    syncFromDom();

    const meta = infraMeta();
    if(summary){
      if(meta.agriM2 == null){
        summary.textContent = 'Enter Agricultural area in Section 2 to calculate % impacted.';
      } else {
        summary.textContent = `Total impacted: ${Math.round(meta.totalImpactM2*100)/100} m² • Agricultural area: ${Math.round(meta.agriM2)} m² • % impacted: ${meta.pct ?? '—'}%`;
      }
    }

    if(status){
      if(meta.agriM2 == null){
        status.textContent = 'Please complete Agricultural area in Section 2.';
        status.classList.add('warn');
      } else {
        status.textContent = '';
        status.classList.remove('warn');
      }
    }

    setCriterionValue('new', 7, meta.pct);

    const next = document.getElementById('s4-next');
    const ok = meta.agriM2 != null;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    validateCriteriaSection(false);
    updateReview();
  }

  if(tbody && tbody.children.length === 0) tbody.insertAdjacentHTML('beforeend', infraRowTemplate(uid('infra')));

  if(BOUND.infra){
    recalc();
    return;
  }
  BOUND.infra = true;

  addBtn.onclick = ()=>{
    markSubmitDirty_('new');
    tbody.insertAdjacentHTML('beforeend', infraRowTemplate(uid('infra')));
    recalc();
  };

  tbody.addEventListener('click', (e)=>{
    const btn = e.target.closest('[data-act="del"]');
    if(!btn) return;
    markSubmitDirty_('new');
    btn.closest('tr')?.remove();
    recalc();
  });

  tbody.addEventListener('input', ()=>{
    markSubmitDirty_('new');
    recalc();
  });
  tbody.addEventListener('change', ()=>{
    markSubmitDirty_('new');
    recalc();
  });

  recalc();
}

function bindCrops(){
  const tbody = document.querySelector('#crops-table tbody');
  const addBtn = document.getElementById('crops-add');
  const summary = document.getElementById('crops-summary');
  const cropsBlock = document.getElementById('crops-block');

  function syncFromDom(){
    const rows = [...tbody.querySelectorAll('tr')];
    const data = [];
    for(const tr of rows){
      const crop = (tr.querySelector('[data-k="crop"]').value || '').trim();
      const areaHa = Number(tr.querySelector('[data-k="area_ha"]').value || 0);
      const n = Number(tr.querySelector('[data-k="n_kg_ha"]').value || 0);
      const on = Number(tr.querySelector('[data-k="org_n_kg_ha"]').value || 0);
      const y = Number(tr.querySelector('[data-k="yield_t_ha"]').value || 0);

      data.push({
        id: tr.dataset.id,
        crop,
        area_ha: Number.isFinite(areaHa)?areaHa:null,
        n_kg_ha: Number.isFinite(n)?n:null,
        org_n_kg_ha: Number.isFinite(on)?on:null,
        yield_t_ha: Number.isFinite(y)?y:null
      });
    }
    app.data.crops = data;
  }

  function recalc(){
    syncFromDom();
    if(summary){
      const n = app.data.crops.length;
      summary.textContent = n ? `${n} crop${n===1?'':'s'} added` : 'No crops added yet.';
    }

    const next = document.getElementById('s6-next');
    const ok = app.data.crops.length > 0;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    if(cropsBlock){
      cropsBlock.classList.toggle('missing-block', !ok);
    }

    updateReview();
    updateSubmitNewState();
  }

  if(tbody && tbody.children.length === 0) tbody.insertAdjacentHTML('beforeend', cropRowTemplate(uid('crop')));

  if(BOUND.crops){
    recalc();
    return;
  }
  BOUND.crops = true;

  addBtn.onclick = ()=>{
    markSubmitDirty_('new');
    tbody.insertAdjacentHTML('beforeend', cropRowTemplate(uid('crop')));
    recalc();
  };

  tbody.addEventListener('click', (e)=>{
    const btn = e.target.closest('[data-act="del"]');
    if(!btn) return;
    markSubmitDirty_('new');
    btn.closest('tr')?.remove();
    recalc();
  });

  tbody.addEventListener('input', ()=>{
    markSubmitDirty_('new');
    recalc();
  });
  tbody.addEventListener('change', ()=>{
    markSubmitDirty_('new');
    recalc();
  });

  recalc();
}

// ===============================
// Criteria inputs (new + existing)
// ===============================
function setCriterionValue(mode, cid, value){
  const c = CRITERIA.find(x=> x.id === cid);
  if(!c) return;

  const isYesNo = c.unit === 'yesno';
  const el = document.getElementById(`crit-${mode}-${cid}`);
  if(!el) return;

  if(isYesNo){
    el.checked = (value != null && Number(value) >= 1);
  } else {
    el.value = (value == null ? '' : String(value));
  }

  const lv = (value==null) ? null : levelFor(value, c.thresholds);
  setBadge(mode, cid, lv);
}

function bindCriteriaNew(){
  renderCriteriaCards(document.getElementById('criteria-new'), 'new');

  const map = [
    {cid:2, key:'c2'},
    {cid:3, key:'c3'},
    {cid:4, key:'c4'}
  ];

  for(const {cid, key} of map){
    const el = document.getElementById(`crit-new-${cid}`);
    el?.addEventListener('input', ()=>{
      const v = Number(el.value);
      app.data.criteria_inputs[key] = el.value.trim()==='' ? null : clampPct(v);
      markSubmitDirty_('new');
      validateCriteriaSection(false);
      updateReview();
    });
  }

  [5,6].forEach(cid=>{
    const el = document.getElementById(`crit-new-${cid}`);
    el?.addEventListener('change', ()=>{
      if(cid===5) app.data.criteria_inputs.c5 = !!el.checked;
      if(cid===6) app.data.criteria_inputs.c6 = !!el.checked;
      const v = el.checked ? 1 : 0;
      setBadge('new', cid, levelFor(v, CRITERIA.find(x=>x.id===cid).thresholds));
      markSubmitDirty_('new');
      validateCriteriaSection(false);
      updateReview();
    });
  });

  setCriterionValue('new', 1, calcRotationCriterion1());
  setCriterionValue('new', 7, infraMeta().pct);

  [5,6].forEach(cid=>{
    const v = (cid===5 ? (app.data.criteria_inputs.c5 ? 1 : 0) : (app.data.criteria_inputs.c6 ? 1 : 0));
    setBadge('new', cid, levelFor(v, CRITERIA.find(x=>x.id===cid).thresholds));
  });

  validateCriteriaSection(false);
}

function bindCriteriaExisting(){
  renderCriteriaCards(document.getElementById('criteria-existing'), 'existing');

  for(const c of CRITERIA){
    const el = document.getElementById(`crit-existing-${c.id}`);
    if(!el) continue;

    if(c.unit === 'yesno'){
      el.addEventListener('change', ()=>{
        const v = el.checked ? 1 : 0;
        app.data.mrvCriteria['C'+c.id] = v;
        setBadge('existing', c.id, levelFor(v, c.thresholds));
        markSubmitDirty_('trade2025');
        validateExisting(false);
      });
    } else {
      el.addEventListener('input', ()=>{
        const raw = el.value.trim();
        const v = raw==='' ? null : clampPct(Number(raw));
        app.data.mrvCriteria['C'+c.id] = v;
        setBadge('existing', c.id, v==null ? null : levelFor(v, c.thresholds));
        markSubmitDirty_('trade2025');
        validateExisting(false);
      });
    }
  }

  validateExisting(false);
}

function validateCriteriaSection(showMissing){
  const values = computeCriteriaForNew();
  const required = [1,2,3,4,7];
  const missing = required.filter(k => values[k] == null);

  // Clear/mark missing fields
  if(showMissing){
    for(const k of required){
      const el = document.getElementById(`crit-new-${k}`);
      if(!el) continue;
      // skip calculated
      if(k===1 || k===7) continue;
      setMissing(el, values[k] == null);
    }
  } else {
    // clear missing styling while typing
    for(const k of [2,3,4]) setMissing(document.getElementById(`crit-new-${k}`), false);
  }

  for(const c of CRITERIA){
    const v = values[c.id];
    if(v==null) { setBadge('new', c.id, null); continue; }
    if(c.unit === 'yesno') continue;
    setBadge('new', c.id, levelFor(v, c.thresholds));
  }

  const next = document.getElementById('s5-next');
  const ok = missing.length === 0;
  next.classList.toggle('disabled', !ok);
  setAriaDisabled_(next, !ok);

  const status = document.getElementById('s5-status');
  if(status){
    if(ok){
      status.textContent = '';
      status.classList.remove('err','warn');
    } else {
      status.textContent = `Please complete: ${missing.map(k=>`C${k}`).join(', ')}`;
      status.classList.remove('err');
      status.classList.add('warn');
    }
  }

  updateReview();
  updateSubmitNewState();
}

function validateExisting(showMissing){
  const required = [1,2,3,4,7];
  const missing = required.filter(i => app.data.mrvCriteria['C'+i] == null);

  if(showMissing){
    for(const i of required){
      const el = document.getElementById(`crit-existing-${i}`);
      setMissing(el, app.data.mrvCriteria['C'+i] == null);
    }
  } else {
    required.forEach(i=> setMissing(document.getElementById(`crit-existing-${i}`), false));
  }

  const btn = document.getElementById('submit-existing');
  const ok = missing.length === 0;
  btn.classList.toggle('disabled', !ok);
  setAriaDisabled_(btn, !ok);
  if(ok && app.submitState.trade2025 !== 'loading') syncSubmitState_('trade2025');

  const status = document.getElementById('existing-status');
  if(status){
    if(ok){
      status.textContent = '';
      status.classList.remove('err','warn');
    } else {
      status.textContent = `Please enter: ${missing.map(i=>`C${i}`).join(', ')}`;
      status.classList.remove('err');
      status.classList.add('warn');
    }
  }
}

// ===============================
// Section navigation / validation
// ===============================
function validateSection1(showMissing){
  const nameEl = document.getElementById('s1-applicant');
  const businessEl = document.getElementById('s1-business');
  const emailEl = document.getElementById('s1-email');
  const typeEl = document.getElementById('s1-type');

  const name = nameEl.value.trim();
  const business = businessEl.value.trim();
  const email = emailEl.value.trim();
  const type = typeEl.value;

  const missing = [];
  if(!name) missing.push('Applicant name');
  if(!business) missing.push('Farm business name');
  if(!email || !email.includes('@')) missing.push('Email address');
  if(!type) missing.push('Farm type');

  const ok = missing.length === 0;

  const btn = document.getElementById('s1-continue');
  btn.classList.toggle('disabled', !ok);
  setAriaDisabled_(btn, !ok);

  if(showMissing){
    setMissing(nameEl, !name);
    setMissing(businessEl, !business);
    setMissing(emailEl, !email || !email.includes('@'));
    setMissing(typeEl, !type);
  } else {
    [nameEl,businessEl,emailEl,typeEl].forEach(el=> setMissing(el,false));
  }

  const status = document.getElementById('s1-status');
  if(status){
    if(ok){
      status.textContent = '';
      status.classList.remove('err','warn');
    } else {
      status.textContent = `Please complete: ${missing.join(', ')}`;
      status.classList.remove('err');
      status.classList.add('warn');
    }
  }

  return { ok, name, business, email, type };
}

function bindSection1(){
  ['s1-applicant','s1-business','s1-email','s1-type'].forEach(id=>{
    document.getElementById(id)?.addEventListener('input', ()=> validateSection1(false));
    document.getElementById(id)?.addEventListener('change', ()=> validateSection1(false));
  });

  document.getElementById('s1-reset')?.addEventListener('click', ()=> resetAll());

  document.getElementById('s1-continue')?.addEventListener('click', ()=>{
    const v = validateSection1(true);
    if(!v.ok) return;

    app.data.applicant = { name: v.name, business: v.business, email: v.email, type: v.type };
    app.flow = (v.type === 'new') ? 'new' : 'trade2025';
    app.submitted = false;
    resetSubmitState_();

    markDone('s1');

    if(app.flow === 'new'){
      document.getElementById('new-flow').classList.remove('hide');
      document.getElementById('existing-flow').classList.add('hide');
      initLocksForFlow_();
      setStepper('s2');
      scrollToId('sec-2');
    } else {
      document.getElementById('existing-flow').classList.remove('hide');
      document.getElementById('new-flow').classList.add('hide');
      initLocksForFlow_();
      setStepper('mrv');
      scrollToId('sec-x');
    }

    saveViewState_({ flow: app.flow });
  });

  validateSection1(false);
}

function bindBaseline(){
  const ids = ['s2-agri','s2-arable','s2-perm-crop','s2-pasture','s2-habitat','s2-livestock','s2-fuel'];
  const status = document.getElementById('s2-status');
  const next = document.getElementById('s2-next');

  function sync(showMissing){
    app.data.baseline.agricultural_ha = numVal('s2-agri');
    app.data.baseline.arable_ha = numVal('s2-arable');
    app.data.baseline.perm_cropland_ha = numVal('s2-perm-crop');
    app.data.baseline.perm_pasture_ha = numVal('s2-pasture');
    app.data.baseline.habitat_ha = numVal('s2-habitat');
    app.data.baseline.livestock = numVal('s2-livestock');
    app.data.baseline.fuel_lpy = numVal('s2-fuel');

    const missing = [];
    if(app.data.baseline.agricultural_ha == null) missing.push({id:'s2-agri', label:'Agricultural area'});
    if(app.data.baseline.arable_ha == null) missing.push({id:'s2-arable', label:'Arable cropland'});
    if(app.data.baseline.perm_cropland_ha == null) missing.push({id:'s2-perm-crop', label:'Permanent cropland'});
    if(app.data.baseline.perm_pasture_ha == null) missing.push({id:'s2-pasture', label:'Permanent pasture'});
    if(app.data.baseline.habitat_ha == null) missing.push({id:'s2-habitat', label:'Natural habitat area'});
    if(app.data.baseline.livestock == null) missing.push({id:'s2-livestock', label:'Number of livestock'});
    if(app.data.baseline.fuel_lpy == null) missing.push({id:'s2-fuel', label:'Fuel consumption'});

    const ok = missing.length === 0;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    if(showMissing){
      ids.forEach(id=> setMissing(document.getElementById(id), false));
      missing.forEach(m=> setMissing(document.getElementById(m.id), true));
    } else {
      ids.forEach(id=> setMissing(document.getElementById(id), false));
    }

    if(status){
      if(ok){
        status.textContent = '';
        status.classList.remove('warn','err');
      } else {
        status.textContent = `Please complete: ${missing.map(m=>m.label).join(', ')}`;
        status.classList.remove('err');
        status.classList.add('warn');
      }
    }

    const meta = infraMeta();
    setCriterionValue('new', 7, meta.pct);

    updateReview();
    validateCriteriaSection(false);
  }

  ids.forEach(id=>{
    document.getElementById(id)?.addEventListener('input', ()=>{
      markSubmitDirty_('new');
      sync(false);
    });
    document.getElementById(id)?.addEventListener('change', ()=>{
      markSubmitDirty_('new');
      sync(false);
    });
  });

  next?.addEventListener('click', ()=>{
    sync(true);
    if(next.getAttribute('aria-disabled') === 'true') return;
    markDone('s2');
    setLocked('sec-3', false);
    setStepper('s3');
    scrollToId('sec-3');
  });

  sync(false);
}

function scrollToId(id){
  const el = document.getElementById(id);
  if(!el) return;
  const header = document.querySelector('.header');
  const offset = (header ? header.offsetHeight : 0) + 12;
  const top = el.getBoundingClientRect().top + window.pageYOffset - offset;
  window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
}

function setLocked(id, locked){
  const el = document.getElementById(id);
  if(!el) return;
  el.classList.toggle('locked', !!locked);
}

function initLocksForFlow_(){
  if(app.flow === 'new'){
    setLocked('sec-2', false);
    ['sec-3','sec-4','sec-5','sec-6','sec-7','sec-out-new'].forEach(id=> setLocked(id, true));
  }
  if(app.flow === 'trade2025'){
    setLocked('sec-x', false);
    setLocked('sec-out-existing', true);
  }
}

function applyLocksFromProgress_(){
  initLocksForFlow_();

  if(app.flow === 'new'){
    if(app.done.has('s2')) setLocked('sec-3', false);
    if(app.done.has('s3')) setLocked('sec-4', false);
    if(app.done.has('s4')) setLocked('sec-5', false);
    if(app.done.has('s5')) setLocked('sec-6', false);
    if(app.done.has('s6')) setLocked('sec-7', false);
    if(app.submitted) setLocked('sec-out-new', false);
  }

  if(app.flow === 'trade2025'){
    if(app.submitted) setLocked('sec-out-existing', false);
  }
}

function bindNewNav(){
  document.getElementById('s3-next')?.addEventListener('click', ()=>{
    if(document.getElementById('s3-next').getAttribute('aria-disabled') === 'true'){
      document.getElementById('rot-block')?.classList.add('missing-block');
      return;
    }
    markDone('s3');
    setLocked('sec-4', false);
    setStepper('s4');
    scrollToId('sec-4');
  });

  document.getElementById('s4-next')?.addEventListener('click', ()=>{
    if(document.getElementById('s4-next').getAttribute('aria-disabled') === 'true') return;
    markDone('s4');
    setLocked('sec-5', false);
    setStepper('s5');
    scrollToId('sec-5');
  });

  document.getElementById('s5-next')?.addEventListener('click', ()=>{
    validateCriteriaSection(true);
    if(document.getElementById('s5-next').getAttribute('aria-disabled') === 'true') return;
    markDone('s5');
    setLocked('sec-6', false);
    setStepper('s6');
    scrollToId('sec-6');
  });

  document.getElementById('s6-next')?.addEventListener('click', ()=>{
    if(document.getElementById('s6-next').getAttribute('aria-disabled') === 'true'){
      document.getElementById('crops-block')?.classList.add('missing-block');
      return;
    }
    markDone('s6');
    setLocked('sec-7', false);
    setStepper('s7');
    scrollToId('sec-7');
    updateReview();
    updateSubmitNewState();
  });
}

function updateReview(){
  if(app.flow !== 'new') return;

  const el = document.getElementById('review');
  if(!el) return;

  const b = app.data.baseline;
  const meta = infraMeta();
  const criteria = computeCriteriaForNew();
  const { levels, overall } = selectEligibleMeasures(criteria);

  el.innerHTML = `
    <div class="grid">
      <div class="card" style="box-shadow:none">
        <div class="section-title">Applicant</div>
        <div><strong>${escapeHtml(app.data.applicant.name)}</strong></div>
        <div class="small">${escapeHtml(app.data.applicant.business)} • ${escapeHtml(app.data.applicant.email)}</div>
      </div>
      <div class="card" style="box-shadow:none">
        <div class="section-title">Baseline</div>
        <div class="small">Agricultural area: <strong>${b.agricultural_ha ?? '—'}</strong> ha</div>
        <div class="small">Arable cropland: <strong>${b.arable_ha ?? '—'}</strong> ha</div>
        <div class="small">Nature infrastructure impact: <strong>${meta.pct ?? '—'}%</strong></div>
      </div>
      <div class="card" style="box-shadow:none">
        <div class="section-title">Current pathway</div>
        <div class="small">Overall level is the lowest of your criteria levels.</div>
        <div class="small">Overall: <strong>${overall ? overall[0].toUpperCase()+overall.slice(1) : '—'}</strong></div>
      </div>
    </div>

    <div class="divider"></div>
    <div class="section-title">Calculated criteria</div>
    ${criteriaSummaryTable(criteria, levels)}

    <div class="divider"></div>
    <div class="section-title">Your entries</div>

    <div class="small"><strong>Rotations:</strong> ${app.data.rotations.length} • <strong>Nature infra entries:</strong> ${app.data.infrastructure.length} • <strong>Crops:</strong> ${app.data.crops.length}</div>
  `;
}

function updateSubmitNewState(){
  if(app.flow !== 'new') return;

  const btn = document.getElementById('submit-new');
  if(!btn) return;
  if(app.submitState.new === 'loading') return;

  const c = computeCriteriaForNew();
  const required = [1,2,3,4,7];
  const missing = required.filter(k => c[k] == null);

  const ok = missing.length === 0 && (app.data.crops.length > 0) && (app.data.rotations.length > 0);
  btn.classList.toggle('disabled', !ok);
  setAriaDisabled_(btn, !ok);
  if(ok) syncSubmitState_('new');
}

// ===============================
// Submission payload & posting
// ===============================
function buildSubmissionPayload(flow, criteriaValues, selection){
  const answers = {};
  const levels = {};

  for(const c of CRITERIA){
    const v = criteriaValues[c.id];
    const key = `C${c.id}:${c.name.replaceAll('Integrated nutrient management principles','Nutrient principles').replaceAll('Integrated pest management principles','IPM principles')}`;
    answers[key] = (v==null ? null : v);

    const lv = (v==null ? null : levelFor(v, c.thresholds));
    levels[`C${c.id}`] = lv;
  }

  const overall = selection.overall;

  return {
    timestamp: new Date().toISOString(),
    farmer: {
      business_name: app.data.applicant.business,
      name: app.data.applicant.name,
      email: app.data.applicant.email,
      consent: true
    },
    overall,
    answers,
    levels,
    measures: selection.measures.map(m=>({
      code:m.code,
      name:m.name,
      group:m.group,
      category:m.category,
      contributes_to:m.contributes_to
    })),
    meta: {
      app: 'Application Form 2026 POC',
      farm_type: flow,
      details: {
        baseline: app.data.baseline,
        rotations: app.data.rotations,
        infrastructure: app.data.infrastructure,
        crops: app.data.crops,
        nature_infra_summary: infraMeta(),
        criteria_values: criteriaValues,
        ineligible_measures: (selection.ineligible || []).map(m=>({ code:m.code, name:m.name, group:m.group, category:m.category }))
      }
    }
  };
}

async function postSubmission(payload){
  const controller = new AbortController();
  const timeout = setTimeout(()=>controller.abort(), 15000);
  try{
    const res = await fetch(SUBMIT_ENDPOINT + '?action=submit', {
      method:'POST',
      headers:{'Content-Type':'text/plain'},
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
    if(!res.ok) throw new Error(`Submission failed (HTTP ${res.status})`);
  } finally {
    clearTimeout(timeout);
  }
}
// ===============================
// Print (new + existing)
// ===============================
function printNewFarm_(selection, criteriaValues, levels){
  const area = document.getElementById('print-area');
  if(!area) return;

  const a = app.data.applicant;
  const b = app.data.baseline;
  const infra = infraMeta();

  const yesNo = (v)=> (v==null ? '—' : (Number(v)>=1 ? 'Yes' : 'No'));
  const fmt = (v, suf='')=> (v==null ? '—' : `${v}${suf}`);

  const rotationsRows = (app.data.rotations || []).map(r=>`<tr><td>${escapeHtml(r.name||'')}</td><td>${escapeHtml(r.num_crops ?? '')}</td><td>${escapeHtml(r.area_ha ?? '')}</td></tr>`).join('');
  const infraRows = (app.data.infrastructure || []).map(r=>`<tr><td>${escapeHtml(r.type||'')}</td><td>${escapeHtml(r.qty ?? '')}</td><td>${escapeHtml(r.unit ?? '')}</td><td>${escapeHtml(r.impact ?? '')}</td><td>${escapeHtml(r.impact_m2 ?? '')}</td></tr>`).join('');
  const cropsRows = (app.data.crops || []).map(r=>`<tr><td>${escapeHtml(r.crop||'')}</td><td>${escapeHtml(r.area_ha ?? '')}</td><td>${escapeHtml(r.n_kg_ha ?? '')}</td><td>${escapeHtml(r.org_n_kg_ha ?? '')}</td><td>${escapeHtml(r.yield_t_ha ?? '')}</td></tr>`).join('');

  const critRows = CRITERIA.map(c=>{
    const v = criteriaValues[c.id];
    const disp = (c.unit === 'yesno') ? yesNo(v) : fmt(v, c.unit==='%'?'%':'');
    const lv = levels[c.id] ? (levels[c.id][0].toUpperCase()+levels[c.id].slice(1)) : '—';
    return `<tr><td class="mono">C${c.id}</td><td>${escapeHtml(c.name)}</td><td>${escapeHtml(String(disp))}</td><td>${escapeHtml(lv)}</td></tr>`;
  }).join('');

  // Eligible & ineligible lists
  const elig = selection.measures || [];
  const inelig = selection.ineligible || [];

  const renderPrintMeasures = (title, list, isIneligible)=>{
    const grouped = groupMeasures(list);
    const groupOrder = ['in-field','capital','resilience'];

    let html = `<h3 style="margin:14px 0 8px">${escapeHtml(title)}</h3>`;
    if(isIneligible){
      html += `<div class="small" style="margin-bottom:8px"><strong>Ineligible:</strong> you cannot apply for these measures.</div>`;
    }

    for(const g of groupOrder){
      const titleG = (g === 'in-field') ? 'In-field' : (g[0].toUpperCase()+g.slice(1));
      const catMap = grouped.get(g);
      html += `<div style="margin-top:10px"><div class="section-title" style="margin:0 0 6px">${escapeHtml(titleG)}</div>`;
      if(!catMap || catMap.size===0){
        html += `<div class="small">None</div></div>`;
        continue;
      }
      const cats = [...catMap.keys()].sort((a,b)=> a.localeCompare(b));
      for(const cat of cats){
        html += `<div style="margin:6px 0"><div style="font-weight:800">${escapeHtml(cat)}</div>`;
        for(const m of catMap.get(cat)){
          html += `<div style="margin:4px 0; ${isIneligible?'opacity:.75':''}"><span class="mono">${escapeHtml(m.code)}</span> — ${escapeHtml(m.name || '')}${isIneligible ? ' (INELIGIBLE)' : ''}</div>`;
        }
        html += `</div>`;
      }
      html += `</div>`;
    }
    return html;
  };

  area.innerHTML = `
    <div style="padding:18px">
      <div class="print-title">Application Form 2026 — Submission Summary</div>
      <div class="print-sub">Generated from your responses in the application form.</div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Applicant &amp; Farm</h3>
      <div><strong>Applicant:</strong> ${escapeHtml(a.name)}</div>
      <div><strong>Farm business:</strong> ${escapeHtml(a.business)}</div>
      <div><strong>Email:</strong> ${escapeHtml(a.email)}</div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Baseline (Section 2)</h3>
      <div class="small">Agricultural area: <strong>${fmt(b.agricultural_ha)}</strong> ha</div>
      <div class="small">Arable cropland: <strong>${fmt(b.arable_ha)}</strong> ha</div>
      <div class="small">Permanent cropland: <strong>${fmt(b.perm_cropland_ha)}</strong> ha</div>
      <div class="small">Permanent pasture: <strong>${fmt(b.perm_pasture_ha)}</strong> ha</div>
      <div class="small">Natural habitat area: <strong>${fmt(b.habitat_ha)}</strong> ha</div>
      <div class="small">Livestock: <strong>${fmt(b.livestock)}</strong></div>
      <div class="small">Fuel: <strong>${fmt(b.fuel_lpy)}</strong> litres/year</div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Rotations (Section 3)</h3>
      <table class="table">
        <thead><tr><th>Rotation</th><th># crops</th><th>Area (ha)</th></tr></thead>
        <tbody>${rotationsRows || `<tr><td colspan="3" class="small">None</td></tr>`}</tbody>
      </table>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Nature infrastructure (Section 4)</h3>
      <div class="small">Total impacted: <strong>${fmt(Math.round(infra.totalImpactM2*100)/100)}</strong> m² • % of agricultural area: <strong>${fmt(infra.pct, '%')}</strong></div>
      <table class="table">
        <thead><tr><th>Type</th><th>Qty</th><th>Unit</th><th>Impact factor</th><th>Impact area (m²)</th></tr></thead>
        <tbody>${infraRows || `<tr><td colspan="5" class="small">None</td></tr>`}</tbody>
      </table>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Criteria (Section 5)</h3>
      <table class="table">
        <thead><tr><th>Criterion</th><th>What it measures</th><th>Your value</th><th>Level</th></tr></thead>
        <tbody>${critRows}</tbody>
      </table>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Crops (Section 6)</h3>
      <table class="table">
        <thead><tr><th>Crop</th><th>Area (ha)</th><th>N (kg/ha)</th><th>Organic N (kg/ha)</th><th>Yield (t/ha)</th></tr></thead>
        <tbody>${cropsRows || `<tr><td colspan="5" class="small">None</td></tr>`}</tbody>
      </table>

      <div class="divider"></div>
      ${renderPrintMeasures('Eligible measures', elig, false)}

      <div class="divider"></div>
      ${renderPrintMeasures('Ineligible measures (for transparency)', inelig, true)}

      <div class="divider"></div>
      <div class="small">Reminder: Return to <strong>NatureBid</strong> and submit proposals for <strong>eligible measures only</strong>.</div>
    </div>
  `;

  // Print
  document.body.classList.add('printing');
  window.print();
  // Cleanup
  setTimeout(()=>{
    document.body.classList.remove('printing');
  }, 300);
}

function printExistingFarm_(selection, criteriaValues, levels){
  const area = document.getElementById('print-area');
  if(!area) return;

  const a = app.data.applicant;
  const yesNo = (v)=> (v==null ? '—' : (Number(v)>=1 ? 'Yes' : 'No'));
  const fmt = (v, suf='')=> (v==null ? '—' : `${v}${suf}`);

  const critRows = CRITERIA.map(c=>{
    const v = criteriaValues[c.id];
    const disp = (c.unit === 'yesno') ? yesNo(v) : fmt(v, c.unit==='%'?'%':'');
    const lv = levels[c.id] ? (levels[c.id][0].toUpperCase()+levels[c.id].slice(1)) : '—';
    return `<tr><td class="mono">C${c.id}</td><td>${escapeHtml(c.name)}</td><td>${escapeHtml(String(disp))}</td><td>${escapeHtml(lv)}</td></tr>`;
  }).join('');

  const elig = selection.measures || [];
  const inelig = selection.ineligible || [];

  const renderPrintMeasures = (title, list, isIneligible)=>{
    const grouped = groupMeasures(list);
    const groupOrder = ['in-field','capital','resilience'];

    let html = `<h3 style="margin:14px 0 8px">${escapeHtml(title)}</h3>`;
    if(isIneligible){
      html += `<div class="small" style="margin-bottom:8px"><strong>Ineligible:</strong> you cannot apply for these measures.</div>`;
    }

    for(const g of groupOrder){
      const titleG = (g === 'in-field') ? 'In-field' : (g[0].toUpperCase()+g.slice(1));
      const catMap = grouped.get(g);
      html += `<div style="margin-top:10px"><div class="section-title" style="margin:0 0 6px">${escapeHtml(titleG)}</div>`;
      if(!catMap || catMap.size===0){
        html += `<div class="small">None</div></div>`;
        continue;
      }
      const cats = [...catMap.keys()].sort((a,b)=> a.localeCompare(b));
      for(const cat of cats){
        html += `<div style="margin:6px 0"><div style="font-weight:800">${escapeHtml(cat)}</div>`;
        for(const m of catMap.get(cat)){
          html += `<div style="margin:4px 0; ${isIneligible?'opacity:.75':''}"><span class="mono">${escapeHtml(m.code)}</span> — ${escapeHtml(m.name || '')}${isIneligible ? ' (INELIGIBLE)' : ''}</div>`;
        }
        html += `</div>`;
      }
      html += `</div>`;
    }
    return html;
  };

  area.innerHTML = `
    <div style="padding:18px">
      <div class="print-title">Application Form 2026 — Submission Summary</div>
      <div class="print-sub">Generated from your MRV criteria values.</div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Applicant &amp; Farm</h3>
      <div><strong>Applicant:</strong> ${escapeHtml(a.name)}</div>
      <div><strong>Farm business:</strong> ${escapeHtml(a.business)}</div>
      <div><strong>Email:</strong> ${escapeHtml(a.email)}</div>
      <div class="small" style="margin-top:6px">Farm type: <strong>Existing Trade 2025 farm</strong></div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Criteria (MRV)</h3>
      <table class="table">
        <thead><tr><th>Criterion</th><th>What it measures</th><th>Your value</th><th>Level</th></tr></thead>
        <tbody>${critRows}</tbody>
      </table>

      <div class="divider"></div>
      ${renderPrintMeasures('Eligible measures', elig, false)}

      <div class="divider"></div>
      ${renderPrintMeasures('Ineligible measures (for transparency)', inelig, true)}

      <div class="divider"></div>
      <div class="small">Reminder: Return to <strong>NatureBid</strong> and submit proposals for <strong>eligible measures only</strong>.</div>
    </div>
  `;

  document.body.classList.add('printing');
  window.print();
  setTimeout(()=>{ document.body.classList.remove('printing'); }, 300);
}

function bindPrintButton_(){
  document.getElementById('print-new')?.addEventListener('click', ()=>{
    if(app.flow !== 'new' || !app.submitted) return;
    const criteria = computeCriteriaForNew();
    const selection = selectEligibleMeasures(criteria);
    printNewFarm_(selection, criteria, selection.levels);
  });

  document.getElementById('print-existing')?.addEventListener('click', ()=>{
    if(app.flow !== 'trade2025' || !app.submitted) return;
    const criteria = computeCriteriaForExisting();
    const selection = selectEligibleMeasures(criteria);
    printExistingFarm_(selection, criteria, selection.levels);
  });
}

// ===============================
// Existing MRV flow submit
// ===============================
function bindExistingSubmit(){
  const btn = document.getElementById('submit-existing');
  const hint = document.getElementById('measures-existing-hint');
  const out = document.getElementById('measures-existing');
  const status = document.getElementById('existing-status');

  btn?.addEventListener('click', async ()=>{
    validateExisting(true);
    if(btn.getAttribute('aria-disabled') === 'true'){
      updateStepperStatusesExisting_(true);
      return;
    }

    setSubmitState_('trade2025', 'loading');
    btn.classList.add('disabled');
    setAriaDisabled_(btn, true);
    if(status){
      status.textContent = '';
      status.classList.remove('err');
    }

    // Default C5/C6 if not provided
    for(let i=5;i<=6;i++){
      const k = 'C'+i;
      if(app.data.mrvCriteria[k] == null) app.data.mrvCriteria[k] = 0;
    }

    // Always try to ensure measures are loaded so we can show eligible + ineligible lists
    const okMeasures = await ensureMeasuresLoaded_();
    if(!okMeasures){
      setSubmitState_('trade2025', app.submitDirty.trade2025 ? 'resubmit' : 'ready');
      btn.classList.remove('disabled');
      setAriaDisabled_(btn, false);
      out.innerHTML = `<div class="error"><strong>Measures could not be loaded</strong><div class="small" style="margin-top:6px">Please refresh the page and try again.</div></div>`;
      return;
    }

    const criteria = computeCriteriaForExisting();
    const selection = selectEligibleMeasures(criteria);

    try{
      const payload = buildSubmissionPayload('trade2025', criteria, selection);
      await postSubmission(payload);

      setSubmitState_('trade2025', 'submitted');
      app.submitDirty.trade2025 = false;
      btn.classList.remove('disabled');
      setAriaDisabled_(btn, false);
      if(status){
        status.textContent = '';
        status.classList.remove('err');
      }

      // Enable print for existing farms
      document.getElementById('print-existing')?.classList.remove('hide');

      app.submitted = true;
      markAllDoneForFlow_();
      setLocked('sec-out-existing', false);
      setStepper('mrv');

      hint.textContent = 'These are the measures you can apply for in NatureBid.';
      renderMeasuresThreeColumns(out, selection.measures);

      // Ineligible measures (collapsed by default)
      const inelEl = document.getElementById('ineligible-existing');
      renderMeasuresThreeColumnsIneligible(inelEl, selection.ineligible || []);
      const acc = document.getElementById('ineligible-acc-existing');
      acc?.classList.remove('hide');
      acc?.removeAttribute('open');

      scrollToId('sec-out-existing');

    } catch (e) {
      console.error(e);
      app.submitDirty.trade2025 = true;
      setSubmitState_('trade2025', 'resubmit');
      btn.classList.remove('disabled');
      setAriaDisabled_(btn, false);
      if(status){
        status.textContent = 'Submission failed - please try again.';
        status.classList.add('err');
      }
    }
  });
}

// ===============================
// New flow submit
// ===============================
function bindNewSubmit(){
  const btn = document.getElementById('submit-new');
  const hint = document.getElementById('measures-new-hint');
  const out = document.getElementById('measures-new');
  const status = document.getElementById('submit-status');
  const printBtn = document.getElementById('print-new');

  btn?.addEventListener('click', async ()=>{
    // On submit attempt, show missing markers where applicable
    validateCriteriaSection(true);

    if(btn.getAttribute('aria-disabled') === 'true'){
      updateStepperStatusesNew_(true);

      document.getElementById('crops-block')?.classList.toggle('missing-block', app.data.crops.length === 0);
      document.getElementById('rot-block')?.classList.toggle('missing-block', app.data.rotations.length === 0);
      return;
    }

    setSubmitState_('new', 'loading');
    btn.classList.add('disabled');
    setAriaDisabled_(btn, true);
    if(status){
      status.textContent = '';
      status.classList.remove('err');
    }

    // Always ensure measures loaded so we can show eligible + ineligible lists
    const okMeasures = await ensureMeasuresLoaded_();
    if(!okMeasures){
      setSubmitState_('new', app.submitDirty.new ? 'resubmit' : 'ready');
      btn.classList.remove('disabled');
      setAriaDisabled_(btn, false);
      out.innerHTML = `<div class="error"><div class="loading"><span class="spinner"></span><strong>Measures could not be loaded</strong></div><div class="small" style="margin-top:6px">Please refresh the page and try again.</div></div>`;
      return;
    }

    const criteria = computeCriteriaForNew();
    const selection = selectEligibleMeasures(criteria);

    try{
      const payload = buildSubmissionPayload('new', criteria, selection);
      await postSubmission(payload);

      setSubmitState_('new', 'submitted');
      app.submitDirty.new = false;
      btn.classList.remove('disabled');
      setAriaDisabled_(btn, false);
      if(status){
        status.textContent = '';
        status.classList.remove('err');
      }
      app.submitted = true;
      markAllDoneForFlow_();
      setLocked('sec-out-new', false);
      setStepper('s7');

      hint.textContent = 'These are the measures you can apply for in NatureBid.';
      renderMeasuresThreeColumns(out, selection.measures);

      // Ineligible measures (collapsed by default)
      const inelEl = document.getElementById('ineligible-new');
      renderMeasuresThreeColumnsIneligible(inelEl, selection.ineligible || []);
      const acc = document.getElementById('ineligible-acc-new');
      acc?.classList.remove('hide');
      acc?.removeAttribute('open');

      // Show print button (new farms only, after submit)
      printBtn?.classList.remove('hide');

      scrollToId('sec-out-new');

    } catch (e) {
      console.error(e);
      app.submitDirty.new = true;
      setSubmitState_('new', 'resubmit');
      btn.classList.remove('disabled');
      setAriaDisabled_(btn, false);
      if(status){
        status.textContent = 'Submission failed - please try again.';
        status.classList.add('err');
      }
    }
  });
}

// ===============================
// Reset
// ===============================
function clearAutosave_(){
  try{ localStorage.removeItem(LS_AUTOSAVE_KEY); } catch {}
}

function resetAll(){
  clearMissingWithin(document.body);
  clearAutosave_();
  try{ localStorage.removeItem(LS_VIEW_KEY); } catch {}

  ['s1-applicant','s1-business','s1-email','s1-type'].forEach(id=>{
    const el = document.getElementById(id);
    if(!el) return;
    if(el.tagName === 'SELECT') el.value = '';
    else el.value = '';
  });

  ['s2-agri','s2-arable','s2-perm-crop','s2-pasture','s2-habitat','s2-livestock','s2-fuel'].forEach(id=>{
    const el = document.getElementById(id); if(el) el.value='';
  });

  document.querySelector('#rot-table tbody')?.replaceChildren();
  document.querySelector('#infra-table tbody')?.replaceChildren();
  document.querySelector('#crops-table tbody')?.replaceChildren();

  app.data.criteria_inputs = { c2:null, c3:null, c4:null, c5:false, c6:false };
  app.data.mrvCriteria = {};

  app.flow = null;
  app.submitted = false;
  resetSubmitState_();
  app.done = new Set();
  app.warn = new Set();
  app.data.applicant = { name:'', business:'', email:'', type:'' };
  app.data.baseline = { agricultural_ha:null, arable_ha:null, perm_cropland_ha:null, perm_pasture_ha:null, habitat_ha:null, livestock:null, fuel_lpy:null };
  app.data.rotations = [];
  app.data.infrastructure = [];
  app.data.crops = [];

  document.getElementById('new-flow')?.classList.add('hide');
  document.getElementById('existing-flow')?.classList.add('hide');

  if(document.getElementById('measures-new')) document.getElementById('measures-new').innerHTML='';
  if(document.getElementById('measures-existing')) document.getElementById('measures-existing').innerHTML='';
  document.getElementById('ineligible-acc-new')?.classList.add('hide');
  document.getElementById('ineligible-acc-existing')?.classList.add('hide');
  document.getElementById('print-existing')?.classList.add('hide');

  document.getElementById('print-new')?.classList.add('hide');
  document.getElementById('print-existing')?.classList.add('hide');

  if(document.getElementById('measures-new-hint')) document.getElementById('measures-new-hint').textContent='Submit responses to see eligible measures.';
  if(document.getElementById('measures-existing-hint')) document.getElementById('measures-existing-hint').textContent='Enter your MRV criteria values and click submit to see eligible measures.';

  ['s1-status','s2-status','s3-status','s4-status','s5-status','s6-status','submit-status','existing-status'].forEach(id=>{
    const el = document.getElementById(id); if(el){ el.textContent=''; el.classList.remove('err','warn'); }
  });

  // Reset binding guards
  BOUND.rot = false;
  BOUND.infra = false;
  BOUND.crops = false;

  renderCriteriaCards(document.getElementById('criteria-new'), 'new');
  renderCriteriaCards(document.getElementById('criteria-existing'), 'existing');

  bindRotations();
  bindInfrastructure();
  bindCrops();

  bindCriteriaNew();
  bindCriteriaExisting();

  validateSection1(false);
  setStepper('s1');
}

// ===============================
// View state (minimal)
// ===============================
function saveViewState_(partial){
  try{
    const cur = readViewState_();
    const next = { ...cur, ...partial, _ts: Date.now() };
    localStorage.setItem(LS_VIEW_KEY, JSON.stringify(next));
  } catch {}
}

function readViewState_(){
  try{ return JSON.parse(localStorage.getItem(LS_VIEW_KEY) || '{}') || {}; }
  catch{ return {}; }
}

function readDomAutosaveState_(){
  const state = {
    flow: app.flow,
    submitted: !!app.submitted,
    submitDirty: app.submitDirty,
    done: Array.from(app.done || []),
    applicant: {
      name: document.getElementById('s1-applicant')?.value?.trim() || '',
      business: document.getElementById('s1-business')?.value?.trim() || '',
      email: document.getElementById('s1-email')?.value?.trim() || '',
      type: document.getElementById('s1-type')?.value || ''
    },
    baseline: {
      agricultural_ha: numVal('s2-agri'),
      arable_ha: numVal('s2-arable'),
      perm_cropland_ha: numVal('s2-perm-crop'),
      perm_pasture_ha: numVal('s2-pasture'),
      habitat_ha: numVal('s2-habitat'),
      livestock: numVal('s2-livestock'),
      fuel_lpy: numVal('s2-fuel')
    },
    rotations: [...document.querySelectorAll('#rot-table tbody tr')].map(tr=>({
      id: tr.dataset.id || uid('rot'),
      name: tr.querySelector('[data-k="name"]')?.value?.trim() || '',
      num_crops: (()=>{ const v = Number(tr.querySelector('[data-k="num_crops"]')?.value||''); return Number.isFinite(v)?v:null; })(),
      area_ha: (()=>{ const v = Number(tr.querySelector('[data-k="area_ha"]')?.value||''); return Number.isFinite(v)?v:null; })()
    })),
    infrastructure: [...document.querySelectorAll('#infra-table tbody tr')].map(tr=>({
      id: tr.dataset.id || uid('infra'),
      type: tr.querySelector('[data-k="type"]')?.value || '',
      qty: (()=>{ const v = Number(tr.querySelector('[data-k="qty"]')?.value||''); return Number.isFinite(v)?v:null; })()
    })),
    criteria_inputs: {
      c2: clampPct(numVal('crit-new-2')),
      c3: clampPct(numVal('crit-new-3')),
      c4: clampPct(numVal('crit-new-4')),
      c5: !!document.getElementById('crit-new-5')?.checked,
      c6: !!document.getElementById('crit-new-6')?.checked
    },
    crops: [...document.querySelectorAll('#crops-table tbody tr')].map(tr=>({
      id: tr.dataset.id || uid('crop'),
      crop: tr.querySelector('[data-k="crop"]')?.value?.trim() || '',
      area_ha: (()=>{ const v = Number(tr.querySelector('[data-k="area_ha"]')?.value||''); return Number.isFinite(v)?v:null; })(),
      n_kg_ha: (()=>{ const v = Number(tr.querySelector('[data-k="n_kg_ha"]')?.value||''); return Number.isFinite(v)?v:null; })(),
      org_n_kg_ha: (()=>{ const v = Number(tr.querySelector('[data-k="org_n_kg_ha"]')?.value||''); return Number.isFinite(v)?v:null; })(),
      yield_t_ha: (()=>{ const v = Number(tr.querySelector('[data-k="yield_t_ha"]')?.value||''); return Number.isFinite(v)?v:null; })()
    })),
    mrvCriteria: (function(){
      const out = {};
      for(let i=1;i<=7;i++){
        const c = CRITERIA.find(x=>x.id===i);
        const el = document.getElementById(`crit-existing-${i}`);
        if(!el || !c) continue;
        if(c.unit === 'yesno') out['C'+i] = el.checked ? 1 : 0;
        else {
          const raw = el.value.trim();
          out['C'+i] = raw==='' ? null : clampPct(Number(raw));
        }
      }
      return out;
    })()
  };
  return state;
}

function saveAutosave_(){
  try{
    const state = readDomAutosaveState_();
    localStorage.setItem(LS_AUTOSAVE_KEY, JSON.stringify({ ts: Date.now(), state }));
  } catch {}
}

function scheduleAutosave_(){
  try{
    if(autosaveTimer_) clearTimeout(autosaveTimer_);
    autosaveTimer_ = setTimeout(saveAutosave_, AUTOSAVE_DEBOUNCE_MS);
  } catch {}
}

function readAutosave_(){
  try{
    const raw = localStorage.getItem(LS_AUTOSAVE_KEY);
    if(!raw) return null;
    const obj = JSON.parse(raw);
    if(!obj || !obj.state) return null;
    return obj.state;
  } catch { return null; }
}

function restoreAutosaveToDom_(state){
  if(!state) return false;

  // Section 1
  if(state.applicant){
    const a = state.applicant;
    if(document.getElementById('s1-applicant')) document.getElementById('s1-applicant').value = a.name || '';
    if(document.getElementById('s1-business')) document.getElementById('s1-business').value = a.business || '';
    if(document.getElementById('s1-email')) document.getElementById('s1-email').value = a.email || '';
    if(document.getElementById('s1-type')) document.getElementById('s1-type').value = a.type || '';

    // set app flow from restored type
    if(a.type === 'new') app.flow = 'new';
    else if(a.type === 'trade2025') app.flow = 'trade2025';
  }

  // Baseline
  if(state.baseline){
    const b = state.baseline;
    const map = [
      ['s2-agri','agricultural_ha'],
      ['s2-arable','arable_ha'],
      ['s2-perm-crop','perm_cropland_ha'],
      ['s2-pasture','perm_pasture_ha'],
      ['s2-habitat','habitat_ha'],
      ['s2-livestock','livestock'],
      ['s2-fuel','fuel_lpy']
    ];
    for(const [id,k] of map){
      const el = document.getElementById(id);
      if(!el) continue;
      el.value = (b[k] == null ? '' : String(b[k]));
    }
  }

  // Rotations table
  if(Array.isArray(state.rotations)){
    const tbody = document.querySelector('#rot-table tbody');
    if(tbody){
      tbody.replaceChildren();
      const rows = state.rotations.length ? state.rotations : [{id:uid('rot'), name:'', num_crops:null, area_ha:null}];
      for(const r of rows){
        tbody.insertAdjacentHTML('beforeend', rotRowTemplate(r.id || uid('rot')));
        const tr = tbody.querySelector(`tr[data-id="${r.id}"]`) || tbody.lastElementChild;
        if(tr){
          tr.querySelector('[data-k="name"]').value = r.name || '';
          tr.querySelector('[data-k="num_crops"]').value = (r.num_crops == null ? '' : String(r.num_crops));
          tr.querySelector('[data-k="area_ha"]').value = (r.area_ha == null ? '' : String(r.area_ha));
        }
      }
    }
  }

  // Infrastructure table
  if(Array.isArray(state.infrastructure)){
    const tbody = document.querySelector('#infra-table tbody');
    if(tbody){
      tbody.replaceChildren();
      const rows = state.infrastructure.length ? state.infrastructure : [{id:uid('infra'), type:'', qty:null}];
      for(const r of rows){
        tbody.insertAdjacentHTML('beforeend', infraRowTemplate(r.id || uid('infra')));
        const tr = tbody.querySelector(`tr[data-id="${r.id}"]`) || tbody.lastElementChild;
        if(tr){
          tr.querySelector('[data-k="type"]').value = r.type || '';
          tr.querySelector('[data-k="qty"]').value = (r.qty == null ? '' : String(r.qty));
        }
      }
    }
  }

  // Crops table
  if(Array.isArray(state.crops)){
    const tbody = document.querySelector('#crops-table tbody');
    if(tbody){
      tbody.replaceChildren();
      const rows = state.crops.length ? state.crops : [{id:uid('crop'), crop:'', area_ha:null, n_kg_ha:null, org_n_kg_ha:null, yield_t_ha:null}];
      for(const r of rows){
        tbody.insertAdjacentHTML('beforeend', cropRowTemplate(r.id || uid('crop')));
        const tr = tbody.querySelector(`tr[data-id="${r.id}"]`) || tbody.lastElementChild;
        if(tr){
          tr.querySelector('[data-k="crop"]').value = r.crop || '';
          tr.querySelector('[data-k="area_ha"]').value = (r.area_ha == null ? '' : String(r.area_ha));
          tr.querySelector('[data-k="n_kg_ha"]').value = (r.n_kg_ha == null ? '' : String(r.n_kg_ha));
          tr.querySelector('[data-k="org_n_kg_ha"]').value = (r.org_n_kg_ha == null ? '' : String(r.org_n_kg_ha));
          tr.querySelector('[data-k="yield_t_ha"]').value = (r.yield_t_ha == null ? '' : String(r.yield_t_ha));
        }
      }
    }
  }

  // Stash non-table values; will be applied after criteria cards are rendered
  app.data.criteria_inputs = state.criteria_inputs || app.data.criteria_inputs;
  app.data.mrvCriteria = state.mrvCriteria || app.data.mrvCriteria;
  app.submitted = !!state.submitted;
  app.submitDirty = state.submitDirty || { new:false, trade2025:false };
  app.done = new Set(Array.isArray(state.done) ? state.done : []);

  return true;
}

function applyRestoredCriteriaToDom_(){
  // New criteria
  if(app.flow === 'new'){
    if(document.getElementById('crit-new-2')) document.getElementById('crit-new-2').value = (app.data.criteria_inputs.c2 == null ? '' : String(app.data.criteria_inputs.c2));
    if(document.getElementById('crit-new-3')) document.getElementById('crit-new-3').value = (app.data.criteria_inputs.c3 == null ? '' : String(app.data.criteria_inputs.c3));
    if(document.getElementById('crit-new-4')) document.getElementById('crit-new-4').value = (app.data.criteria_inputs.c4 == null ? '' : String(app.data.criteria_inputs.c4));
    if(document.getElementById('crit-new-5')) document.getElementById('crit-new-5').checked = !!app.data.criteria_inputs.c5;
    if(document.getElementById('crit-new-6')) document.getElementById('crit-new-6').checked = !!app.data.criteria_inputs.c6;
  }

  // Existing MRV criteria
  if(app.flow === 'trade2025'){
    for(let i=1;i<=7;i++){
      const c = CRITERIA.find(x=>x.id===i);
      const el = document.getElementById(`crit-existing-${i}`);
      if(!el || !c) continue;
      const v = app.data.mrvCriteria['C'+i];
      if(c.unit === 'yesno') el.checked = (v != null && Number(v) >= 1);
      else el.value = (v == null ? '' : String(v));
    }
  }
}

// ===============================
// Init
// ===============================
function bindResets(){
  document.getElementById('reset-new')?.addEventListener('click', ()=> resetAll());
  document.getElementById('reset-existing')?.addEventListener('click', ()=>{
    clearMissingWithin(document.getElementById('existing-flow'));
    app.data.mrvCriteria = {};
    app.submitted = false;
    app.submitDirty.trade2025 = false;
    setSubmitState_('trade2025', 'ready');
    for(const c of CRITERIA){
      const el = document.getElementById(`crit-existing-${c.id}`);
      if(!el) continue;
      if(c.unit === 'yesno') el.checked = false;
      else el.value = '';
      setBadge('existing', c.id, null);
    }
    validateExisting(false);
    document.getElementById('measures-existing').innerHTML='';
    document.getElementById('measures-existing-hint').textContent = 'Enter your MRV criteria values and click submit to see eligible measures.';
    document.getElementById('ineligible-acc-existing')?.classList.add('hide');
  });
}

function runTests(){
  try{
    console.assert(typeof LS_VIEW_KEY === 'string' && LS_VIEW_KEY.length > 0, 'LS_VIEW_KEY defined');
    console.assert(document.getElementById('sec-1'), 'Section 1 exists');
    console.assert(document.getElementById('s1-continue'), 'Section 1 continue exists');
    console.assert(Array.isArray(INFRA_LOOKUP) && INFRA_LOOKUP.length === 10, 'Infra lookup has 10 items');
    console.assert(CRITERIA.length === 7, '7 criteria');
    console.assert(typeof bindInfrastructure === 'function' && typeof bindCrops === 'function', 'Bindings exist');
    console.log('All tests passed');
  } catch(e){
    console.warn('Tests failed:', e);
  }
}

async function main(){
  setStepper('s1');

  // Try to apply cached measures immediately (fast UI)
  try{
    const cached = readMeasuresCache_();
    if(cached) applyMeasuresFromSheet_(cached.measures, 'Using cached measures');
  } catch {}

  // Restore saved user progress (silent)
  const restoredState = readAutosave_();
  if(restoredState){
    restoreAutosaveToDom_(restoredState);
  }

  // Always attempt to load measures on startup (farmers should not need to click anything)
  await loadConfigFromSheet(false);

  bindMeasuresModal_();
  bindSection1();

  // If a saved flow exists, show it immediately (no extra click)
  if(app.flow === 'new'){
    document.getElementById('new-flow')?.classList.remove('hide');
    document.getElementById('existing-flow')?.classList.add('hide');
  } else if(app.flow === 'trade2025'){
    document.getElementById('existing-flow')?.classList.remove('hide');
    document.getElementById('new-flow')?.classList.add('hide');
  }
  if(app.flow) applyLocksFromProgress_();

  bindBaseline();
  bindRotations();
  bindInfrastructure();
  bindCriteriaNew();
  bindCrops();
  bindNewNav();
  bindNewSubmit();

  bindCriteriaExisting();
  bindExistingSubmit();

  // Apply restored criteria values after the cards exist
  applyRestoredCriteriaToDom_();
  if(app.flow === 'new') syncSubmitState_('new');
  else if(app.flow === 'trade2025') syncSubmitState_('trade2025');
  else resetSubmitState_();

  bindResets();
  bindPrintButton_();
  runTests();

  // If we restored a flow, compute current step & re-render outputs if previously submitted
  if(app.flow){
    const activeKey = app.submitted ? (app.flow === 'new' ? 's7' : 'mrv') : (app.flow === 'new' ? 's2' : 'mrv');
    setStepper(activeKey);

    if(app.submitted){
      const okMeasures = await ensureMeasuresLoaded_();
      if(okMeasures){
        if(app.flow === 'new'){
          const criteria = computeCriteriaForNew();
          const selection = selectEligibleMeasures(criteria);
          document.getElementById('measures-new-hint').textContent = 'These are the measures you can apply for in NatureBid.';
          renderMeasuresThreeColumns(document.getElementById('measures-new'), selection.measures);
          renderMeasuresThreeColumnsIneligible(document.getElementById('ineligible-new'), selection.ineligible || []);
          document.getElementById('ineligible-acc-new')?.classList.remove('hide');
          document.getElementById('print-new')?.classList.remove('hide');
          document.getElementById('new-flow')?.classList.remove('hide');
          setLocked('sec-out-new', false);
        } else {
          const criteria = computeCriteriaForExisting();
          const selection = selectEligibleMeasures(criteria);
          document.getElementById('measures-existing-hint').textContent = 'These are the measures you can apply for in NatureBid.';
          renderMeasuresThreeColumns(document.getElementById('measures-existing'), selection.measures);
          document.getElementById('print-existing')?.classList.remove('hide');
          renderMeasuresThreeColumnsIneligible(document.getElementById('ineligible-existing'), selection.ineligible || []);
          document.getElementById('ineligible-acc-existing')?.classList.remove('hide');
          document.getElementById('existing-flow')?.classList.remove('hide');
          setLocked('sec-out-existing', false);
        }
      }
    }
  }

  // Silent autosave (no banner)
  document.addEventListener('input', scheduleAutosave_, true);
  document.addEventListener('change', scheduleAutosave_, true);
  scheduleAutosave_();
}

document.addEventListener('DOMContentLoaded', main);
