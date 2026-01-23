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
  // Criteria 5 & 6: calculated from Synthetic inputs
  {id:5,name:'Integrated nutrient management principles',question:'Calculated from the Synthetic inputs section.',thresholds:{leading:3},unit:'principles',calc:'synthetic'},
  {id:6,name:'Integrated pest management principles',question:'Calculated from the Synthetic inputs section.',thresholds:{leading:3},unit:'principles',calc:'synthetic'},
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

const REGION_OPTIONS = [
  { value:'east_of_england', label:'East of England', aggregators:['cefetra','chilton','frontier','openfield'] },
  { value:'yorkshire', label:'Yorkshire', aggregators:['frontier','openfield'] }
];

const AGGREGATOR_LABELS = {
  cefetra: 'Cefetra',
  chilton: 'Chilton',
  frontier: 'Frontier',
  openfield: 'Openfield'
};

const AGGREGATOR_COLUMN_MAP = {
  east_of_england: {
    frontier: 'frontier_eoe',
    openfield: 'openfield_eoe',
    cefetra: 'cefetra',
    chilton: 'chilton'
  },
  yorkshire: {
    frontier: 'frontier_yks',
    openfield: 'openfield_yks'
  }
};

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
  validation: {
    rotations: { hasComplete:false, hasPartial:false },
    infrastructure: { hasComplete:false, hasPartial:false },
    crops: { hasComplete:false, hasPartial:false }
  },
  data: {
    applicant: { name:'', business:'', email:'', type:'', region:'', supply_aggregator:'' },
    water: { anglian:null, affinity:null },
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
    criteria_inputs: { c2:null, c3:null, c4:null },
    synthetic_inputs: {
      c5: { q1:null, q2:null, q3:null },
      c6: { q1:null, q2:null, q3:null }
    },
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

function yesNoToBool(v){
  if(v === true) return true;
  if(v === false) return false;
  const s = String(v ?? '').trim().toLowerCase();
  if(s === 'yes' || s === 'true' || s === '1') return true;
  if(s === 'no' || s === 'false' || s === '0') return false;
  return null;
}

function boolToYesNo(v){
  if(v === true) return 'yes';
  if(v === false) return 'no';
  return '';
}

function normalizeGroup_(g){
  const s = String(g ?? '').toLowerCase().trim();
  if(!s) return 'in-field';
  if(s === 'infield' || s === 'in field' || s === 'in-field' || s === 'in_field') return 'in-field';
  if(s === 'capital') return 'capital';
  if(s === 'resilience' || s === 'resilience payment') return 'resilience';
  return s;
}

function regionLabel_(v){
  return REGION_OPTIONS.find(r=> r.value === v)?.label || v || '';
}

function aggregatorLabel_(v){
  return AGGREGATOR_LABELS[v] || v || '';
}

function aggregatorsForRegion_(regionValue){
  return REGION_OPTIONS.find(r=> r.value === regionValue)?.aggregators || [];
}

function aggregatorColumnForSelection_(regionValue, aggregatorValue){
  if(!regionValue || !aggregatorValue) return null;
  const map = AGGREGATOR_COLUMN_MAP[regionValue];
  return map ? (map[aggregatorValue] || null) : null;
}

function setAggregatorOptions_(regionValue){
  const aggEl = document.getElementById('s1-aggregator');
  if(!aggEl) return;

  const prev = aggEl.value;
  const options = aggregatorsForRegion_(regionValue);
  const placeholder = regionValue ? 'Select…' : 'Select region first…';

  aggEl.innerHTML = `<option value="" selected disabled>${placeholder}</option>` +
    options.map(opt=> `<option value="${opt}">${escapeHtml(aggregatorLabel_(opt))}</option>`).join('');
  aggEl.disabled = !regionValue;

  if(prev && options.includes(prev)){
    aggEl.value = prev;
  } else {
    aggEl.value = '';
  }
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

    const aggregatorFlags = {
      frontier_eoe: isBoolTrue(m.frontier_eoe),
      cefetra: isBoolTrue(m.cefetra),
      openfield_eoe: isBoolTrue(m.openfield_eoe),
      chilton: isBoolTrue(m.chilton),
      openfield_yks: isBoolTrue(m.openfield_yks),
      frontier_yks: isBoolTrue(m.frontier_yks)
    };

    const waterFlags = {
      anglian: isBoolTrue(m.anglian),
      affinity: isBoolTrue(m.affinity)
    };

    cfg[code] = {
      active: isBoolTrue(m.active),
      group: normalizeGroup_(rawGroup),
      category: String(rawCat ?? '').trim(),
      criteria,
      aggregatorFlags,
      waterFlags
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
      criteriaFlags: c.criteria || {},
      aggregatorFlags: c.aggregatorFlags || {},
      waterFlags: c.waterFlags || {}
    });
  }
  out.sort((a,b)=> a.code.localeCompare(b.code));
  return out;
}

function effectiveMeasures(){
  // For eligibility calculations, we only include measures that contribute to at least one criterion.
  return allActiveMeasuresForLists_().filter(m => (m.contributes_to || []).length > 0);
}

function measuresForEligibility_(){
  const all = effectiveMeasures();
  const column = aggregatorColumnForSelection_(app.data.applicant.region, app.data.applicant.supply_aggregator);
  if(!column) return [];
  return all.filter(m => isBoolTrue(m.aggregatorFlags?.[column]));
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
  const order=['below_entry','entry','engaged','advanced','leading'];
  return levels.reduce((m,lv)=> order.indexOf(lv)<order.indexOf(m)?lv:m,'leading');
}

function suggestedMeasures(measures, levelsByCrit){
  if(!measures || measures.length === 0) return [];
  const weak=new Set(Object.entries(levelsByCrit)
    .filter(([,v])=>v==='below_entry'||v==='entry'||v==='engaged')
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
    {key:'s1', label:'Identification'},
    ...(app.flow ? [{key:'s2w', label:'Water eligibility'}] : []),
    ...(app.flow === 'new' ? [
      {key:'s2', label:'Baseline'},
      {key:'s3', label:'Rotations'},
      {key:'s4', label:'Nature infra'},
      {key:'s5', label:'Synthetic inputs'},
      {key:'s6', label:'Criteria'},
      {key:'s7', label:'Crops'},
      {key:'s8', label:'Review'}
    ] : []),
    ...(app.flow === 'trade2025' ? [
      {key:'syn', label:'Synthetic inputs'},
      {key:'mrv', label:'MRV criteria'}
    ] : []),
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
  const waterOk = validateWaterSection(showMissing).ok;
  const baselineReq = ['s2-agri','s2-arable','s2-perm-crop','s2-pasture','s2-habitat','s2-livestock','s2-fuel'];
  const baselineMissing = baselineReq.filter(id => {
    const el = document.getElementById(id);
    return !el || String(el.value||'').trim()==='';
  });

  const arable = app.data.baseline.arable_ha;
  const rotTotal = app.data.rotations.reduce((s,r)=> s + (r.area_ha || 0), 0);
  const rotHasOne = app.validation.rotations?.hasComplete || app.data.rotations.length > 0;
  const rotHasPartial = !!app.validation.rotations?.hasPartial;
  const rotTooMuch = (arable!=null && arable>0 && rotTotal > arable + 1e-9);
  const meta = infraMeta();
  const infraHasOne = app.validation.infrastructure?.hasComplete || app.data.infrastructure.length > 0;
  const infraHasPartial = !!app.validation.infrastructure?.hasPartial;
  const infraOk = meta.agriM2 != null && infraHasOne && !infraHasPartial;

  const syntheticOk = syntheticMeta_().allComplete;
  const c = computeCriteriaForNew();
  const requiredCrit = [1,2,3,4,7];
  const critMissing = requiredCrit.filter(k => c[k] == null);
  if(showMissing){
    validateCriteriaSection(true);
    validateSyntheticSection(true, 'new');
  }

  const cropsHasOne = app.validation.crops?.hasComplete || app.data.crops.length > 0;
  const cropsHasPartial = !!app.validation.crops?.hasPartial;
  const cropsOk = cropsHasOne && !cropsHasPartial;

  const statuses = {
    s1: validateSection1(false).ok,
    s2w: waterOk,
    s2: baselineMissing.length === 0,
    s3: rotHasOne && !rotHasPartial && !rotTooMuch,
    s4: infraOk,
    s5: syntheticOk,
    s6: critMissing.length === 0,
    s7: cropsOk,
    s8: true
  };

  // Set done/warn
  app.warn = new Set();
  for(const k of Object.keys(statuses)){
    if(statuses[k]) app.done.add(k);
    else { app.done.delete(k); app.warn.add(k); }
  }

  setStepper('s8');
  return statuses;
}

function updateStepperStatusesExisting_(showMissing){
  const s1ok = validateSection1(false).ok;
  const waterOk = validateWaterSection(showMissing).ok;
  if(showMissing) validateExisting(true);
  if(showMissing) validateSyntheticSection(true, 'existing');

  const syntheticOk = syntheticMeta_().allComplete;
  const required = [1,2,3,4,7];
  const mrvMissing = required.filter(i => app.data.mrvCriteria['C'+i] == null);
  const mrvOk = mrvMissing.length === 0;

  app.warn = new Set();
  if(s1ok) app.done.add('s1'); else { app.done.delete('s1'); app.warn.add('s1'); }
  if(waterOk) app.done.add('s2w'); else { app.done.delete('s2w'); app.warn.add('s2w'); }
  if(syntheticOk) app.done.add('syn'); else { app.done.delete('syn'); app.warn.add('syn'); }
  if(mrvOk) app.done.add('mrv'); else { app.done.delete('mrv'); app.warn.add('mrv'); }

  setStepper('mrv');
  return { s1: s1ok, s2w: waterOk, syn: syntheticOk, mrv: mrvOk };
}

function markAllDoneForFlow_(){
  app.warn = new Set();
  if(app.flow === 'new') ['s1','s2w','s2','s3','s4','s5','s6','s7','s8'].forEach(k=> app.done.add(k));
  if(app.flow === 'trade2025') ['s1','s2w','syn','mrv'].forEach(k=> app.done.add(k));
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

function isCalculatedCriterion_(mode, c){
  if(c && c.calc === 'synthetic') return true;
  if(mode === 'new' && (c.id === 1 || c.id === 7)) return true;
  return false;
}

function formatLevelLabel_(lv){
  if(!lv) return '—';
  if(lv === 'below_entry') return 'Below entry';
  return lv[0].toUpperCase() + lv.slice(1);
}

// ===============================
// UI rendering (criteria cards)
// ===============================
function renderCriteriaCards(container, mode){
  container.innerHTML='';
  for(const c of CRITERIA){
    const isYesNo = c.unit === 'yesno';
    const isCalculated = isCalculatedCriterion_(mode, c);
    const isSynthetic = c.calc === 'synthetic';
    const unitSuffix = (c.unit === '%') ? '%' : (c.unit === 'principles' ? 'principles' : '');

    const disabledAttr = isCalculated ? 'disabled' : '';

    const inputHtml = isYesNo
      ? `<label class="small" style="display:block;margin-top:10px">
           <input type="checkbox" id="crit-${mode}-${c.id}" ${disabledAttr}>
           Yes
         </label>`
      : `<div class="input-row" style="margin-top:10px">
           <input id="crit-${mode}-${c.id}" type="number" ${disabledAttr} placeholder="${isCalculated?'Calculated':''}" min="0" ${isSynthetic ? 'max="3"' : ''} step="1">
           <span>${unitSuffix}</span>
           <span class="badge" id="badge-${mode}-${c.id}"><span class="dot"></span><span class="level" id="badge-text-${mode}-${c.id}">—</span></span>
         </div>`;

    container.insertAdjacentHTML('beforeend',`
      <div class="card">
        <div class="section-title">Criterion C${c.id}${isCalculated ? ' (calculated)' : ''}</div>
        <h3>${escapeHtml(c.name)}</h3>
        <div class="small">${escapeHtml(c.question)}</div>
        ${isSynthetic ? '' : rangeText(c.thresholds,c.unit)}
        ${inputHtml}
        ${isYesNo ? `<div class="input-row" style="margin-top:10px"><span class="badge" id="badge-${mode}-${c.id}"><span class="dot"></span><span class="level" id="badge-text-${mode}-${c.id}">—</span></span></div>` : ''}
        ${isSynthetic ? `<div class="small" style="margin-top:8px">Leading if all 3 answers are Yes; otherwise below entry.</div>` : ''}
        ${!isSynthetic && isCalculated && mode==='new' ? `<div class="small" style="margin-top:8px">This is calculated from earlier sections.</div>` : ''}
      </div>
    `);
  }
}

function setBadge(mode, cid, lv){
  const badge = document.getElementById(`badge-${mode}-${cid}`);
  const text = document.getElementById(`badge-text-${mode}-${cid}`);
  if(badge && text){
    badge.className = `badge ${lv||''}`;
    text.textContent = formatLevelLabel_(lv);
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

function syntheticMeta_(){
  const s = app.data.synthetic_inputs || {};
  const pack = (key)=>{
    const group = s[key] || {};
    const vals = [group.q1, group.q2, group.q3];
    const complete = vals.every(v => v === true || v === false);
    const yesCount = vals.filter(v => v === true).length;
    const allYes = complete && yesCount === 3;
    return { yesCount, complete, allYes };
  };
  const c5 = pack('c5');
  const c6 = pack('c6');
  return { c5, c6, allComplete: c5.complete && c6.complete };
}

const SYNTHETIC_QUESTIONS = [
  { key:'c5.q1', label:'Integrated nutrient management: Legal compliance' },
  { key:'c5.q2', label:'Integrated nutrient management: Planning' },
  { key:'c5.q3', label:'Integrated nutrient management: Evidence' },
  { key:'c6.q1', label:'Integrated pest management: Legal compliance' },
  { key:'c6.q2', label:'Integrated pest management: Good practice at handling' },
  { key:'c6.q3', label:'Integrated pest management: Planning' }
];

function setSyntheticValue_(key, value){
  const [crit, q] = key.split('.');
  if(!crit || !q) return;
  if(!app.data.synthetic_inputs) app.data.synthetic_inputs = { c5:{}, c6:{} };
  if(!app.data.synthetic_inputs[crit]) app.data.synthetic_inputs[crit] = {};
  app.data.synthetic_inputs[crit][q] = value;
}

function syncSyntheticFromDom_(){
  for(const q of SYNTHETIC_QUESTIONS){
    const el = document.querySelector(`[data-synthetic="${q.key}"]`);
    setSyntheticValue_(q.key, yesNoToBool(el?.value));
  }
}

function updateSyntheticCriteriaDisplays_(){
  const syn = syntheticMeta_();
  setCriterionValue('new', 5, syn.c5.yesCount);
  setCriterionValue('new', 6, syn.c6.yesCount);
  setCriterionValue('existing', 5, syn.c5.yesCount);
  setCriterionValue('existing', 6, syn.c6.yesCount);
}

function validateSyntheticSection(showMissing, flow){
  const meta = syntheticMeta_();
  const ok = meta.allComplete;

  const btnId = (flow === 'existing') ? 'synth-next' : 's5-next';
  const statusId = (flow === 'existing') ? 'synth-status' : 's5-status';
  const btn = document.getElementById(btnId);
  if(btn){
    btn.classList.toggle('disabled', !ok);
    setAriaDisabled_(btn, !ok);
  }

  const status = document.getElementById(statusId);
  if(status){
    if(ok){
      status.textContent = '';
      status.classList.remove('err','warn');
    } else if(showMissing){
      status.textContent = 'Please answer all Synthetic inputs questions.';
      status.classList.remove('err');
      status.classList.add('warn');
    } else {
      status.textContent = '';
      status.classList.remove('err','warn');
    }
  }

  return { ok };
}

function syntheticCriterionLevel_(cid){
  const meta = syntheticMeta_();
  const group = cid === 5 ? meta.c5 : meta.c6;
  if(!group.complete) return null;
  return group.allYes ? 'leading' : 'below_entry';
}

function criterionLevelForValue_(cid, value){
  if(cid === 5 || cid === 6) return syntheticCriterionLevel_(cid);
  const c = CRITERIA.find(x=> x.id === cid);
  if(!c) return null;
  return value == null ? null : levelFor(value, c.thresholds);
}

function computeCriteriaForNew(){
  const c1 = calcRotationCriterion1();
  const c7 = infraMeta().pct;
  const c2 = app.data.criteria_inputs.c2;
  const c3 = app.data.criteria_inputs.c3;
  const c4 = app.data.criteria_inputs.c4;
  const syn = syntheticMeta_();
  const c5 = syn.c5.yesCount;
  const c6 = syn.c6.yesCount;
  return { 1:c1, 2:c2, 3:c3, 4:c4, 5:c5, 6:c6, 7:c7 };
}

function computeCriteriaForExisting(){
  const out = {};
  for(let i=1;i<=7;i++){
    if(i === 5 || i === 6){
      const syn = syntheticMeta_();
      out[i] = (i === 5 ? syn.c5.yesCount : syn.c6.yesCount);
      continue;
    }
    const v = app.data.mrvCriteria['C'+i];
    out[i] = v == null ? null : v;
  }
  return out;
}

function criteriaLevelsFromValues(values){
  const levels = {};
  for(const c of CRITERIA){
    const v = values[c.id];
    const lv = criterionLevelForValue_(c.id, v);
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
function waterOverrideEligible_(measure){
  const water = app.data.water || {};
  const inAnglian = water.anglian === true;
  const inAffinity = water.affinity === true;
  if(!inAnglian && !inAffinity) return false;
  const flags = measure?.waterFlags || {};
  return (inAnglian && isBoolTrue(flags.anglian)) || (inAffinity && isBoolTrue(flags.affinity));
}

function selectEligibleMeasures(criteriaValues){
  const levels = criteriaLevelsFromValues(criteriaValues);
  const levelsForNonResilience = {};
  for(const c of CRITERIA){
    if(c.id === 5 || c.id === 6) continue;
    levelsForNonResilience[c.id] = levels[c.id];
  }
  const overall = overallFromLevels(levelsForNonResilience);
  const resilienceOk = (levels[5] === 'leading' && levels[6] === 'leading');

  // Eligible measures
  let eligible = [];
  const baseCandidates = measuresForEligibility_();
  const allActive = effectiveMeasures();
  const waterEligible = allActive.filter(m=> waterOverrideEligible_(m));
  if(overall === 'advanced' || overall === 'leading'){
    eligible = resilienceOk ? [renderResiliencePayment()] : [];
  } else {
    eligible = suggestedMeasures(baseCandidates, levelsForNonResilience);
  }
  if(waterEligible.length){
    const seen = new Set(eligible.map(m=>m.code));
    for(const m of waterEligible){
      if(!seen.has(m.code)){
        eligible.push(m);
        seen.add(m.code);
      }
    }
  }

  // Ineligible measures = all active measures not in eligible
  const eligibleCodes = new Set(eligible.map(m=>m.code));
  const candidateMap = new Map();
  for(const m of baseCandidates) candidateMap.set(m.code, m);
  for(const m of waterEligible) candidateMap.set(m.code, m);
  const allCandidates = Array.from(candidateMap.values());
  const ineligible = allCandidates.filter(m=> !eligibleCodes.has(m.code));

  return { overall, levels, measures: eligible, ineligible };
}

function criteriaSummaryTable(values, levels){
  const badgeHtml = (lv)=>{
    if(!lv) return '—';
    const label = formatLevelLabel_(lv);
    return `<span class="badge ${lv}"><span class="dot"></span><span class="level">${label}</span></span>`;
  };

  const rows = CRITERIA.map(c=>{
    const v = values[c.id];
    const lv = levels[c.id];
    const disp = (v==null)
      ? '—'
      : (c.unit === 'yesno'
        ? (v>=1 ? 'Yes' : 'No')
        : (c.unit === 'principles' ? `${v} principles` : `${v}${c.unit==='%'?'%':''}`));

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

  function syncFromDom(){
    const rows = [...tbody.querySelectorAll('tr')];
    const data = [];
    let hasComplete = false;
    let hasPartial = false;
    for(const tr of rows){
      const nameEl = tr.querySelector('[data-k="name"]');
      const numEl = tr.querySelector('[data-k="num_crops"]');
      const areaEl = tr.querySelector('[data-k="area_ha"]');

      const name = (nameEl.value || '').trim();
      const numRaw = (numEl.value || '').trim();
      const areaRaw = (areaEl.value || '').trim();
      const numVal = numRaw === '' ? null : Number(numRaw);
      const areaVal = areaRaw === '' ? null : Number(areaRaw);
      const numCrops = Number.isFinite(numVal) ? numVal : null;
      const areaHa = Number.isFinite(areaVal) ? areaVal : null;

      const any = !!name || numRaw !== '' || areaRaw !== '';
      const complete = !!name && numCrops != null && areaHa != null;
      if(any && !complete) hasPartial = true;
      if(complete){
        hasComplete = true;
        data.push({ id: tr.dataset.id, name, num_crops: numCrops, area_ha: areaHa });
      }

    }
    app.data.rotations = data;
    app.validation.rotations = { hasComplete, hasPartial };
    return { hasComplete, hasPartial };
  }

  function recalc(){
    const { hasComplete, hasPartial } = syncFromDom();

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
      } else if(!hasComplete){
        status.textContent = 'Please add at least one completed rotation.';
        status.classList.remove('err');
        status.classList.add('warn');
      } else if(hasPartial){
        status.textContent = 'Please complete all rotation rows (name, # crops, area).';
        status.classList.remove('err');
        status.classList.add('warn');
      } else {
        status.textContent = '';
        status.classList.remove('err','warn');
      }
    }

    const c1 = calcRotationCriterion1();
    setCriterionValue('new', 1, c1);

    const next = document.getElementById('s3-next');
    const ok = hasComplete && !hasPartial && !tooMuch;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    validateCriteriaSection(false);
    updateReview();
    updateSubmitNewState();
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
    let hasComplete = false;
    let hasPartial = false;
    for(const tr of rows){
      const typeEl = tr.querySelector('[data-k="type"]');
      const qtyEl = tr.querySelector('[data-k="qty"]');

      const type = typeEl.value;
      const qtyRaw = (qtyEl.value || '').trim();
      const qtyVal = qtyRaw === '' ? null : Number(qtyRaw);
      const qty = Number.isFinite(qtyVal) ? qtyVal : null;

      const any = !!type || qtyRaw !== '';
      const complete = !!type && qty != null;
      if(any && !complete) hasPartial = true;
      if(complete) hasComplete = true;

      const info = lookup(type);
      const unit = info ? info.unit : null;
      const impact = info ? info.impact : null;
      const impactM2 = (info && qty != null) ? qty * info.impact : null;

      if(complete){
        data.push({ id: tr.dataset.id, type, qty, unit, impact, impact_m2: impactM2 });
      }

      tr.querySelector('[data-k="unit"]').textContent = unit ?? '—';
      tr.querySelector('[data-k="impact"]').textContent = impact ?? '—';
      tr.querySelector('[data-k="impact_m2"]').textContent = (impactM2==null ? '—' : Math.round(impactM2*100)/100);

    }
    app.data.infrastructure = data;
    app.validation.infrastructure = { hasComplete, hasPartial };
    return { hasComplete, hasPartial };
  }

  function recalc(){
    const { hasComplete, hasPartial } = syncFromDom();

    const meta = infraMeta();
    if(summary){
      if(meta.agriM2 == null){
        summary.textContent = 'Enter Agricultural area in Section 3 to calculate % impacted.';
      } else {
        summary.textContent = `Total impacted: ${Math.round(meta.totalImpactM2*100)/100} m² • Agricultural area: ${Math.round(meta.agriM2)} m² • % impacted: ${meta.pct ?? '—'}%`;
      }
    }

    if(status){
      if(meta.agriM2 == null){
        status.textContent = 'Please complete Agricultural area in Section 3.';
        status.classList.add('warn');
        status.classList.remove('err');
      } else if(!hasComplete){
        status.textContent = 'Please add at least one completed nature infrastructure row.';
        status.classList.add('warn');
        status.classList.remove('err');
      } else if(hasPartial){
        status.textContent = 'Please complete all nature infrastructure rows (type, quantity).';
        status.classList.add('warn');
        status.classList.remove('err');
      } else {
        status.textContent = '';
        status.classList.remove('warn','err');
      }
    }

    setCriterionValue('new', 7, meta.pct);

    const next = document.getElementById('s4-next');
    const ok = meta.agriM2 != null && hasComplete && !hasPartial;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    validateCriteriaSection(false);
    updateReview();
    updateSubmitNewState();
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

function bindSyntheticInputs(){
  const fields = [...document.querySelectorAll('[data-synthetic]')];
  if(!fields.length) return;

  function onChange(e){
    const el = e.target.closest('[data-synthetic]');
    if(!el) return;
    const key = el.getAttribute('data-synthetic');
    const value = yesNoToBool(el.value);
    setSyntheticValue_(key, value);
    document.querySelectorAll(`[data-synthetic="${key}"]`).forEach(other=>{
      if(other !== el) other.value = boolToYesNo(value);
    });
    if(app.flow === 'new') markSubmitDirty_('new');
    if(app.flow === 'trade2025') markSubmitDirty_('trade2025');
    updateSyntheticCriteriaDisplays_();
    validateSyntheticSection(false, 'new');
    validateSyntheticSection(false, 'existing');
    validateCriteriaSection(false);
    validateExisting(false);
    updateReview();
    updateSubmitNewState();
  }

  fields.forEach(el=>{
    el.addEventListener('change', onChange);
    el.addEventListener('input', onChange);
  });

  syncSyntheticFromDom_();
  updateSyntheticCriteriaDisplays_();
  validateSyntheticSection(false, 'new');
  validateSyntheticSection(false, 'existing');
}

function bindCrops(){
  const tbody = document.querySelector('#crops-table tbody');
  const addBtn = document.getElementById('crops-add');
  const summary = document.getElementById('crops-summary');
  const status = document.getElementById('s7-status');

  function syncFromDom(){
    const rows = [...tbody.querySelectorAll('tr')];
    const data = [];
    let hasComplete = false;
    let hasPartial = false;
    for(const tr of rows){
      const cropEl = tr.querySelector('[data-k="crop"]');
      const areaEl = tr.querySelector('[data-k="area_ha"]');
      const nEl = tr.querySelector('[data-k="n_kg_ha"]');
      const onEl = tr.querySelector('[data-k="org_n_kg_ha"]');
      const yEl = tr.querySelector('[data-k="yield_t_ha"]');

      const crop = (cropEl.value || '').trim();
      const areaRaw = (areaEl.value || '').trim();
      const nRaw = (nEl.value || '').trim();
      const onRaw = (onEl.value || '').trim();
      const yRaw = (yEl.value || '').trim();

      const areaVal = areaRaw === '' ? null : Number(areaRaw);
      const nVal = nRaw === '' ? null : Number(nRaw);
      const onVal = onRaw === '' ? null : Number(onRaw);
      const yVal = yRaw === '' ? null : Number(yRaw);

      const areaHa = Number.isFinite(areaVal) ? areaVal : null;
      const n = Number.isFinite(nVal) ? nVal : null;
      const on = Number.isFinite(onVal) ? onVal : null;
      const y = Number.isFinite(yVal) ? yVal : null;

      const any = !!crop || areaRaw !== '' || nRaw !== '' || onRaw !== '' || yRaw !== '';
      const complete = !!crop && areaHa != null && n != null && on != null && y != null;
      if(any && !complete) hasPartial = true;
      if(complete){
        hasComplete = true;
        data.push({
          id: tr.dataset.id,
          crop,
          area_ha: areaHa,
          n_kg_ha: n,
          org_n_kg_ha: on,
          yield_t_ha: y
        });
      }

    }
    app.data.crops = data;
    app.validation.crops = { hasComplete, hasPartial };
    return { hasComplete, hasPartial };
  }

  function recalc(){
    const { hasComplete, hasPartial } = syncFromDom();
    if(summary){
      const n = app.data.crops.length;
      summary.textContent = n ? `${n} crop${n===1?'':'s'} added` : 'No crops added yet.';
    }

    const next = document.getElementById('s7-next');
    const ok = hasComplete && !hasPartial;
    next.classList.toggle('disabled', !ok);
    setAriaDisabled_(next, !ok);

    if(status){
      if(!hasComplete){
        status.textContent = 'Please add at least one completed crop row.';
        status.classList.add('warn');
        status.classList.remove('err');
      } else if(hasPartial){
        status.textContent = 'Please complete all crop rows (crop, area, N, organic N, yield).';
        status.classList.add('warn');
        status.classList.remove('err');
      } else {
        status.textContent = '';
        status.classList.remove('warn','err');
      }
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

  const lv = criterionLevelForValue_(cid, value);
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

  setCriterionValue('new', 1, calcRotationCriterion1());
  setCriterionValue('new', 7, infraMeta().pct);
  const syn = syntheticMeta_();
  setCriterionValue('new', 5, syn.c5.yesCount);
  setCriterionValue('new', 6, syn.c6.yesCount);

  validateCriteriaSection(false);
}

function bindCriteriaExisting(){
  renderCriteriaCards(document.getElementById('criteria-existing'), 'existing');

  for(const c of CRITERIA){
    const el = document.getElementById(`crit-existing-${c.id}`);
    if(!el) continue;

    if(c.calc === 'synthetic'){
      const syn = syntheticMeta_();
      const value = (c.id === 5 ? syn.c5.yesCount : syn.c6.yesCount);
      setCriterionValue('existing', c.id, value);
      continue;
    }

    if(c.unit === 'yesno'){
      el.addEventListener('change', ()=>{
        const v = el.checked ? 1 : 0;
        app.data.mrvCriteria['C'+c.id] = v;
        setBadge('existing', c.id, criterionLevelForValue_(c.id, v));
        markSubmitDirty_('trade2025');
        validateExisting(false);
      });
    } else {
      el.addEventListener('input', ()=>{
        const raw = el.value.trim();
        const v = raw==='' ? null : clampPct(Number(raw));
        app.data.mrvCriteria['C'+c.id] = v;
        setBadge('existing', c.id, v==null ? null : criterionLevelForValue_(c.id, v));
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

  for(const c of CRITERIA){
    const v = values[c.id];
    if(v==null) { setBadge('new', c.id, null); continue; }
    if(c.unit === 'yesno') continue;
    setBadge('new', c.id, criterionLevelForValue_(c.id, v));
  }

  const next = document.getElementById('s6-next');
  const ok = missing.length === 0;
  next.classList.toggle('disabled', !ok);
  setAriaDisabled_(next, !ok);

  const status = document.getElementById('s6-status');
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
  const waterOk = app.data.water && app.data.water.anglian != null && app.data.water.affinity != null;
  const syntheticOk = syntheticMeta_().allComplete;

  const btn = document.getElementById('submit-existing');
  const ok = missing.length === 0 && waterOk && syntheticOk;
  btn.classList.toggle('disabled', !ok);
  setAriaDisabled_(btn, !ok);
  if(ok && app.submitState.trade2025 !== 'loading') syncSubmitState_('trade2025');

  const status = document.getElementById('existing-status');
  if(status){
    if(ok){
      status.textContent = '';
      status.classList.remove('err','warn');
    } else if(!waterOk){
      status.textContent = 'Please complete Water catchment eligibility (Section 2).';
      status.classList.remove('err');
      status.classList.add('warn');
    } else if(!syntheticOk){
      status.textContent = 'Please complete Synthetic inputs (Section 3).';
      status.classList.remove('err');
      status.classList.add('warn');
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
  const regionEl = document.getElementById('s1-region');
  const aggregatorEl = document.getElementById('s1-aggregator');

  const name = nameEl.value.trim();
  const business = businessEl.value.trim();
  const email = emailEl.value.trim();
  const type = typeEl.value;
  const region = regionEl.value;
  const supply_aggregator = aggregatorEl.value;

  const missing = [];
  if(!name) missing.push('Applicant name');
  if(!business) missing.push('Farm business name');
  if(!email || !email.includes('@')) missing.push('Email address');
  if(!type) missing.push('Farm type');
  if(!region) missing.push('Region');
  if(!supply_aggregator) missing.push('Supply aggregator');

  const ok = missing.length === 0;

  const btn = document.getElementById('s1-continue');
  btn.classList.toggle('disabled', !ok);
  setAriaDisabled_(btn, !ok);

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

  if(app.data && app.data.applicant){
    app.data.applicant = { ...app.data.applicant, name, business, email, type, region, supply_aggregator };
  }

  return { ok, name, business, email, type, region, supply_aggregator };
}

function validateWaterSection(showMissing){
  const anglianEl = document.getElementById('s2-anglian');
  const affinityEl = document.getElementById('s2-affinity');
  const anglian = yesNoToBool(anglianEl?.value);
  const affinity = yesNoToBool(affinityEl?.value);

  const missing = [];
  if(anglian == null) missing.push('Anglian Water catchment');
  if(affinity == null) missing.push('Affinity Water catchment');

  const ok = missing.length === 0;

  const btn = document.getElementById('s2w-next');
  if(btn){
    btn.classList.toggle('disabled', !ok);
    setAriaDisabled_(btn, !ok);
  }

  const status = document.getElementById('s2w-status');
  if(status){
    if(ok){
      status.textContent = '';
      status.classList.remove('err','warn');
    } else if(showMissing){
      status.textContent = `Please complete: ${missing.join(', ')}`;
      status.classList.remove('err');
      status.classList.add('warn');
    } else {
      status.textContent = '';
      status.classList.remove('err','warn');
    }
  }

  if(app.data){
    app.data.water = { anglian, affinity };
  }

  return { ok, anglian, affinity };
}

function bindSection1(){
  function markSection1Dirty_(){
    if(app.flow === 'new') markSubmitDirty_('new');
    if(app.flow === 'trade2025') markSubmitDirty_('trade2025');
  }

  ['s1-applicant','s1-business','s1-email','s1-type'].forEach(id=>{
    document.getElementById(id)?.addEventListener('input', ()=>{
      validateSection1(false);
      updateReview();
      markSection1Dirty_();
    });
    document.getElementById(id)?.addEventListener('change', ()=>{
      validateSection1(false);
      updateReview();
      markSection1Dirty_();
    });
  });

  document.getElementById('s1-reset')?.addEventListener('click', ()=> resetAll());

  const regionEl = document.getElementById('s1-region');
  const aggregatorEl = document.getElementById('s1-aggregator');

  regionEl?.addEventListener('change', ()=>{
    const prev = aggregatorEl?.value || '';
    setAggregatorOptions_(regionEl.value);
    const next = aggregatorEl?.value || '';
    if(prev && prev !== next){
      validateSection1(true);
    } else {
      validateSection1(false);
    }
    updateReview();
    markSection1Dirty_();
  });

  aggregatorEl?.addEventListener('change', ()=>{
    validateSection1(false);
    updateReview();
    markSection1Dirty_();
  });

  setAggregatorOptions_(regionEl?.value || '');

  document.getElementById('s1-continue')?.addEventListener('click', ()=>{
    const v = validateSection1(true);
    if(!v.ok) return;

    app.data.applicant = {
      name: v.name,
      business: v.business,
      email: v.email,
      type: v.type,
      region: v.region,
      supply_aggregator: v.supply_aggregator
    };
    app.flow = (v.type === 'new') ? 'new' : 'trade2025';
    app.submitted = false;
    resetSubmitState_();

    markDone('s1');

    if(app.flow === 'new'){
      document.getElementById('sec-2')?.classList.remove('hide');
      document.getElementById('new-flow').classList.remove('hide');
      document.getElementById('existing-flow').classList.add('hide');
      initLocksForFlow_();
      setStepper('s2w');
      scrollToId('sec-2');
    } else {
      document.getElementById('sec-2')?.classList.remove('hide');
      document.getElementById('existing-flow').classList.remove('hide');
      document.getElementById('new-flow').classList.add('hide');
      initLocksForFlow_();
      setStepper('s2w');
      scrollToId('sec-2');
    }

    saveViewState_({ flow: app.flow });
  });

  validateSection1(false);
}

function bindWaterSection(){
  function markWaterDirty_(){
    if(app.flow === 'new') markSubmitDirty_('new');
    if(app.flow === 'trade2025') markSubmitDirty_('trade2025');
  }

  ['s2-anglian','s2-affinity'].forEach(id=>{
    document.getElementById(id)?.addEventListener('change', ()=>{
      validateWaterSection(false);
      updateReview();
      if(app.flow === 'new') updateSubmitNewState();
      if(app.flow === 'trade2025') validateExisting(false);
      markWaterDirty_();
    });
    document.getElementById(id)?.addEventListener('input', ()=>{
      validateWaterSection(false);
      updateReview();
      if(app.flow === 'new') updateSubmitNewState();
      if(app.flow === 'trade2025') validateExisting(false);
      markWaterDirty_();
    });
  });

  document.getElementById('s2w-next')?.addEventListener('click', ()=>{
    const v = validateWaterSection(true);
    if(!v.ok) return;
    markDone('s2w');
    if(app.flow === 'new'){
      setLocked('sec-3', false);
      setStepper('s2');
      scrollToId('sec-3');
    } else if(app.flow === 'trade2025'){
      setLocked('sec-synth', false);
      setStepper('syn');
      scrollToId('sec-synth');
    }
  });

  validateWaterSection(false);
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
    setLocked('sec-4', false);
    setStepper('s3');
    scrollToId('sec-4');
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
    ['sec-3','sec-4','sec-5','sec-6','sec-7','sec-8','sec-9','sec-out-new'].forEach(id=> setLocked(id, true));
  }
  if(app.flow === 'trade2025'){
    setLocked('sec-2', false);
    setLocked('sec-synth', true);
    setLocked('sec-x', true);
    setLocked('sec-out-existing', true);
  }
}

function applyLocksFromProgress_(){
  initLocksForFlow_();

  if(app.flow === 'new'){
    if(app.done.has('s2w')) setLocked('sec-3', false);
    if(app.done.has('s2')) setLocked('sec-4', false);
    if(app.done.has('s3')) setLocked('sec-5', false);
    if(app.done.has('s4')) setLocked('sec-6', false);
    if(app.done.has('s5')) setLocked('sec-7', false);
    if(app.done.has('s6')) setLocked('sec-8', false);
    if(app.done.has('s7')) setLocked('sec-9', false);
    if(app.submitted) setLocked('sec-out-new', false);
  }

  if(app.flow === 'trade2025'){
    if(app.done.has('s2w')) setLocked('sec-synth', false);
    if(app.done.has('syn')) setLocked('sec-x', false);
    if(app.submitted) setLocked('sec-out-existing', false);
  }
}

function bindNewNav(){
  document.getElementById('s3-next')?.addEventListener('click', ()=>{
    if(document.getElementById('s3-next').getAttribute('aria-disabled') === 'true'){
      return;
    }
    markDone('s3');
    setLocked('sec-5', false);
    setStepper('s4');
    scrollToId('sec-5');
  });

  document.getElementById('s4-next')?.addEventListener('click', ()=>{
    if(document.getElementById('s4-next').getAttribute('aria-disabled') === 'true'){
      return;
    }
    markDone('s4');
    setLocked('sec-6', false);
    setStepper('s5');
    scrollToId('sec-6');
  });

  document.getElementById('s5-next')?.addEventListener('click', ()=>{
    validateSyntheticSection(true, 'new');
    if(document.getElementById('s5-next').getAttribute('aria-disabled') === 'true') return;
    markDone('s5');
    setLocked('sec-7', false);
    setStepper('s6');
    scrollToId('sec-7');
  });

  document.getElementById('s6-next')?.addEventListener('click', ()=>{
    validateCriteriaSection(true);
    if(document.getElementById('s6-next').getAttribute('aria-disabled') === 'true') return;
    markDone('s6');
    setLocked('sec-8', false);
    setStepper('s7');
    scrollToId('sec-8');
  });

  document.getElementById('s7-next')?.addEventListener('click', ()=>{
    if(document.getElementById('s7-next').getAttribute('aria-disabled') === 'true'){
      return;
    }
    markDone('s7');
    setLocked('sec-9', false);
    setStepper('s8');
    scrollToId('sec-9');
    updateReview();
    updateSubmitNewState();
  });
}

function bindSyntheticNav(){
  document.getElementById('synth-next')?.addEventListener('click', ()=>{
    validateSyntheticSection(true, 'existing');
    if(document.getElementById('synth-next').getAttribute('aria-disabled') === 'true') return;
    markDone('syn');
    setLocked('sec-x', false);
    setStepper('mrv');
    scrollToId('sec-x');
  });
}

function updateReview(){
  if(app.flow !== 'new') return;

  const el = document.getElementById('review');
  if(!el) return;

  const b = app.data.baseline;
  const w = app.data.water || {};
  const meta = infraMeta();
  const criteria = computeCriteriaForNew();
  const { levels, overall } = selectEligibleMeasures(criteria);
  const waterLabel = (v)=> (v == null ? '—' : (v ? 'Yes' : 'No'));

  el.innerHTML = `
    <div class="grid">
        <div class="card" style="box-shadow:none">
          <div class="section-title">Applicant</div>
          <div><strong>${escapeHtml(app.data.applicant.name)}</strong></div>
          <div class="small">${escapeHtml(app.data.applicant.business)} • ${escapeHtml(app.data.applicant.email)}</div>
          <div class="small">${escapeHtml(regionLabel_(app.data.applicant.region))} • ${escapeHtml(aggregatorLabel_(app.data.applicant.supply_aggregator))}</div>
          <div class="small">Anglian Water catchment: <strong>${waterLabel(w.anglian)}</strong> • Affinity Water catchment: <strong>${waterLabel(w.affinity)}</strong></div>
        </div>
      <div class="card" style="box-shadow:none">
        <div class="section-title">Baseline</div>
        <div class="small">Agricultural area: <strong>${b.agricultural_ha ?? '—'}</strong> ha</div>
        <div class="small">Arable cropland: <strong>${b.arable_ha ?? '—'}</strong> ha</div>
        <div class="small">Nature infrastructure impact: <strong>${meta.pct ?? '—'}%</strong></div>
      </div>
      <div class="card" style="box-shadow:none">
        <div class="section-title">Current pathway</div>
        <div class="small">Overall level is based on Criteria 1–4 and 7. Criteria 5–6 only affect the Resilience Payment.</div>
        <div class="small">Overall: <strong>${formatLevelLabel_(overall)}</strong></div>
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

  const rotOk = app.validation.rotations?.hasComplete && !app.validation.rotations?.hasPartial;
  const infraOk = app.validation.infrastructure?.hasComplete && !app.validation.infrastructure?.hasPartial && infraMeta().agriM2 != null;
  const cropsOk = app.validation.crops?.hasComplete && !app.validation.crops?.hasPartial;
  const waterOk = app.data.water && app.data.water.anglian != null && app.data.water.affinity != null;
  const syntheticOk = syntheticMeta_().allComplete;
  const ok = missing.length === 0 && rotOk && infraOk && cropsOk && waterOk && syntheticOk;
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

    levels[`C${c.id}`] = criterionLevelForValue_(c.id, v);
  }

  const overall = selection.overall;

  return {
    timestamp: new Date().toISOString(),
    region: app.data.applicant.region || null,
    supply_aggregator: app.data.applicant.supply_aggregator || null,
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
        water_catchment: app.data.water,
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
  const w = app.data.water || {};
  const infra = infraMeta();

  const yesNo = (v)=> (v==null ? '—' : (Number(v)>=1 ? 'Yes' : 'No'));
  const fmt = (v, suf='')=> (v==null ? '—' : `${v}${suf}`);

  const rotationsRows = (app.data.rotations || []).map(r=>`<tr><td>${escapeHtml(r.name||'')}</td><td>${escapeHtml(r.num_crops ?? '')}</td><td>${escapeHtml(r.area_ha ?? '')}</td></tr>`).join('');
  const infraRows = (app.data.infrastructure || []).map(r=>`<tr><td>${escapeHtml(r.type||'')}</td><td>${escapeHtml(r.qty ?? '')}</td><td>${escapeHtml(r.unit ?? '')}</td><td>${escapeHtml(r.impact ?? '')}</td><td>${escapeHtml(r.impact_m2 ?? '')}</td></tr>`).join('');
  const cropsRows = (app.data.crops || []).map(r=>`<tr><td>${escapeHtml(r.crop||'')}</td><td>${escapeHtml(r.area_ha ?? '')}</td><td>${escapeHtml(r.n_kg_ha ?? '')}</td><td>${escapeHtml(r.org_n_kg_ha ?? '')}</td><td>${escapeHtml(r.yield_t_ha ?? '')}</td></tr>`).join('');

  const critRows = CRITERIA.map(c=>{
    const v = criteriaValues[c.id];
    const disp = (c.unit === 'yesno')
      ? yesNo(v)
      : (c.unit === 'principles' ? (v==null ? '—' : `${v} principles`) : fmt(v, c.unit==='%'?'%':''));
    const lv = formatLevelLabel_(levels[c.id]);
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
        <div><strong>Region:</strong> ${escapeHtml(regionLabel_(a.region))}</div>
        <div><strong>Supply aggregator:</strong> ${escapeHtml(aggregatorLabel_(a.supply_aggregator))}</div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Water catchment eligibility (Section 2)</h3>
      <div class="small">Anglian Water catchment: <strong>${yesNo(w.anglian)}</strong></div>
      <div class="small">Affinity Water catchment: <strong>${yesNo(w.affinity)}</strong></div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Baseline (Section 3)</h3>
      <div class="small">Agricultural area: <strong>${fmt(b.agricultural_ha)}</strong> ha</div>
      <div class="small">Arable cropland: <strong>${fmt(b.arable_ha)}</strong> ha</div>
      <div class="small">Permanent cropland: <strong>${fmt(b.perm_cropland_ha)}</strong> ha</div>
      <div class="small">Permanent pasture: <strong>${fmt(b.perm_pasture_ha)}</strong> ha</div>
      <div class="small">Natural habitat area: <strong>${fmt(b.habitat_ha)}</strong> ha</div>
      <div class="small">Livestock: <strong>${fmt(b.livestock)}</strong></div>
      <div class="small">Fuel: <strong>${fmt(b.fuel_lpy)}</strong> litres/year</div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Rotations (Section 4)</h3>
      <table class="table">
        <thead><tr><th>Rotation</th><th># crops</th><th>Area (ha)</th></tr></thead>
        <tbody>${rotationsRows || `<tr><td colspan="3" class="small">None</td></tr>`}</tbody>
      </table>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Nature infrastructure (Section 5)</h3>
      <div class="small">Total impacted: <strong>${fmt(Math.round(infra.totalImpactM2*100)/100)}</strong> m² • % of agricultural area: <strong>${fmt(infra.pct, '%')}</strong></div>
      <table class="table">
        <thead><tr><th>Type</th><th>Qty</th><th>Unit</th><th>Impact factor</th><th>Impact area (m²)</th></tr></thead>
        <tbody>${infraRows || `<tr><td colspan="5" class="small">None</td></tr>`}</tbody>
      </table>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Criteria (Section 6)</h3>
      <table class="table">
        <thead><tr><th>Criterion</th><th>What it measures</th><th>Your value</th><th>Level</th></tr></thead>
        <tbody>${critRows}</tbody>
      </table>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Crops (Section 7)</h3>
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
  const w = app.data.water || {};
  const yesNo = (v)=> (v==null ? '—' : (Number(v)>=1 ? 'Yes' : 'No'));
  const fmt = (v, suf='')=> (v==null ? '—' : `${v}${suf}`);

  const critRows = CRITERIA.map(c=>{
    const v = criteriaValues[c.id];
    const disp = (c.unit === 'yesno')
      ? yesNo(v)
      : (c.unit === 'principles' ? (v==null ? '—' : `${v} principles`) : fmt(v, c.unit==='%'?'%':''));
    const lv = formatLevelLabel_(levels[c.id]);
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
        <div><strong>Region:</strong> ${escapeHtml(regionLabel_(a.region))}</div>
        <div><strong>Supply aggregator:</strong> ${escapeHtml(aggregatorLabel_(a.supply_aggregator))}</div>
        <div class="small" style="margin-top:6px">Farm type: <strong>Existing Trade 2025 farm</strong></div>

      <div class="divider"></div>

      <h3 style="margin:0 0 8px">Water catchment eligibility (Section 2)</h3>
      <div class="small">Anglian Water catchment: <strong>${yesNo(w.anglian)}</strong></div>
      <div class="small">Affinity Water catchment: <strong>${yesNo(w.affinity)}</strong></div>

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
    validateSyntheticSection(true, 'existing');
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
    validateSyntheticSection(true, 'new');

    if(btn.getAttribute('aria-disabled') === 'true'){
      updateStepperStatusesNew_(true);
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
      setStepper('s8');

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
  clearAutosave_();
  try{ localStorage.removeItem(LS_VIEW_KEY); } catch {}

  ['s1-applicant','s1-business','s1-email','s1-type','s1-region','s1-aggregator'].forEach(id=>{
    const el = document.getElementById(id);
    if(!el) return;
    if(el.tagName === 'SELECT') el.value = '';
    else el.value = '';
  });
  setAggregatorOptions_('');

  ['s2-anglian','s2-affinity'].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.value = '';
  });

  document.querySelectorAll('[data-synthetic]')?.forEach(el=>{
    el.value = '';
  });

  ['s2-agri','s2-arable','s2-perm-crop','s2-pasture','s2-habitat','s2-livestock','s2-fuel'].forEach(id=>{
    const el = document.getElementById(id); if(el) el.value='';
  });

  document.querySelector('#rot-table tbody')?.replaceChildren();
  document.querySelector('#infra-table tbody')?.replaceChildren();
  document.querySelector('#crops-table tbody')?.replaceChildren();

  app.data.criteria_inputs = { c2:null, c3:null, c4:null };
  app.data.synthetic_inputs = {
    c5: { q1:null, q2:null, q3:null },
    c6: { q1:null, q2:null, q3:null }
  };
  app.data.mrvCriteria = {};

  app.flow = null;
  app.submitted = false;
  resetSubmitState_();
  app.done = new Set();
  app.warn = new Set();
  app.validation = {
    rotations: { hasComplete:false, hasPartial:false },
    infrastructure: { hasComplete:false, hasPartial:false },
    crops: { hasComplete:false, hasPartial:false }
  };
  app.data.applicant = { name:'', business:'', email:'', type:'', region:'', supply_aggregator:'' };
  app.data.water = { anglian:null, affinity:null };
  app.data.baseline = { agricultural_ha:null, arable_ha:null, perm_cropland_ha:null, perm_pasture_ha:null, habitat_ha:null, livestock:null, fuel_lpy:null };
  app.data.rotations = [];
  app.data.infrastructure = [];
  app.data.crops = [];

  document.getElementById('sec-2')?.classList.add('hide');
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

  ['s1-status','s2w-status','s2-status','s3-status','s4-status','s5-status','s6-status','s7-status','synth-status','submit-status','existing-status'].forEach(id=>{
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
  updateSyntheticCriteriaDisplays_();
  validateSyntheticSection(false, 'new');
  validateSyntheticSection(false, 'existing');

  validateSection1(false);
  validateWaterSection(false);
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
        type: document.getElementById('s1-type')?.value || '',
        region: document.getElementById('s1-region')?.value || '',
        supply_aggregator: document.getElementById('s1-aggregator')?.value || ''
      },
    water: {
      anglian: yesNoToBool(document.getElementById('s2-anglian')?.value),
      affinity: yesNoToBool(document.getElementById('s2-affinity')?.value)
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
      c4: clampPct(numVal('crit-new-4'))
    },
    synthetic_inputs: (function(){
      const out = { c5:{}, c6:{} };
      for(const q of SYNTHETIC_QUESTIONS){
        const el = document.querySelector(`[data-synthetic="${q.key}"]`);
        setSyntheticValue_(q.key, yesNoToBool(el?.value));
      }
      out.c5 = { ...(app.data.synthetic_inputs?.c5 || {}) };
      out.c6 = { ...(app.data.synthetic_inputs?.c6 || {}) };
      return out;
    })(),
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
        if(!el || !c || c.calc === 'synthetic') continue;
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
    if(document.getElementById('s1-region')) document.getElementById('s1-region').value = a.region || '';
    setAggregatorOptions_(a.region || '');
    if(document.getElementById('s1-aggregator')) document.getElementById('s1-aggregator').value = a.supply_aggregator || '';

    // set app flow from restored type
    if(a.type === 'new') app.flow = 'new';
    else if(a.type === 'trade2025') app.flow = 'trade2025';
  }

  if(state.water){
    const w = state.water;
    if(document.getElementById('s2-anglian')) document.getElementById('s2-anglian').value = boolToYesNo(w.anglian);
    if(document.getElementById('s2-affinity')) document.getElementById('s2-affinity').value = boolToYesNo(w.affinity);
    app.data.water = { anglian: yesNoToBool(w.anglian), affinity: yesNoToBool(w.affinity) };
  }

  if(state.synthetic_inputs){
    app.data.synthetic_inputs = state.synthetic_inputs;
    for(const q of SYNTHETIC_QUESTIONS){
      const [crit, key] = q.key.split('.');
      const value = state.synthetic_inputs?.[crit]?.[key] ?? null;
      document.querySelectorAll(`[data-synthetic="${q.key}"]`).forEach(el=>{
        el.value = boolToYesNo(value);
      });
    }
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
  app.data.synthetic_inputs = state.synthetic_inputs || app.data.synthetic_inputs;
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
  }

  // Existing MRV criteria
  if(app.flow === 'trade2025'){
    for(let i=1;i<=7;i++){
      const c = CRITERIA.find(x=>x.id===i);
      const el = document.getElementById(`crit-existing-${i}`);
      if(!el || !c) continue;
      const v = app.data.mrvCriteria['C'+i];
      if(c.calc === 'synthetic') continue;
      if(c.unit === 'yesno') el.checked = (v != null && Number(v) >= 1);
      else el.value = (v == null ? '' : String(v));
    }
  }

  updateSyntheticCriteriaDisplays_();
}

// ===============================
// Init
// ===============================
function bindResets(){
  document.getElementById('reset-new')?.addEventListener('click', ()=> resetAll());
  document.getElementById('reset-existing')?.addEventListener('click', ()=>{
    app.data.mrvCriteria = {};
    app.data.synthetic_inputs = {
      c5: { q1:null, q2:null, q3:null },
      c6: { q1:null, q2:null, q3:null }
    };
    document.querySelectorAll('[data-synthetic]')?.forEach(el=>{
      el.value = '';
    });
    app.submitted = false;
    app.submitDirty.trade2025 = false;
    setSubmitState_('trade2025', 'ready');
    for(const c of CRITERIA){
      const el = document.getElementById(`crit-existing-${c.id}`);
      if(!el) continue;
      if(c.calc === 'synthetic'){
        el.value = '';
      } else if(c.unit === 'yesno') {
        el.checked = false;
      } else {
        el.value = '';
      }
      setBadge('existing', c.id, null);
    }
    updateSyntheticCriteriaDisplays_();
    validateSyntheticSection(false, 'existing');
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
      console.assert(document.getElementById('sec-2'), 'Section 2 exists');
      console.assert(document.getElementById('s1-continue'), 'Section 1 continue exists');
      console.assert(document.getElementById('s1-region'), 'Section 1 region exists');
      console.assert(document.getElementById('s1-aggregator'), 'Section 1 aggregator exists');
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
  let usedCachedMeasures = false;
  try{
    const cached = readMeasuresCache_();
    if(cached){
      applyMeasuresFromSheet_(cached.measures, 'Using cached measures');
      usedCachedMeasures = true;
    }
  } catch {}

  // Restore saved user progress (silent)
  const restoredState = readAutosave_();
  if(restoredState){
    restoreAutosaveToDom_(restoredState);
  }

  // Always attempt to load measures on startup (farmers should not need to click anything)
  // If we used cache, refresh in the background so updates in the sheet are picked up.
  if(usedCachedMeasures){
    loadConfigFromSheet(true);
  } else {
    await loadConfigFromSheet(true);
  }

  bindMeasuresModal_();
  bindSection1();
  bindWaterSection();

  // If a saved flow exists, show it immediately (no extra click)
  if(app.flow === 'new'){
    document.getElementById('sec-2')?.classList.remove('hide');
    document.getElementById('new-flow')?.classList.remove('hide');
    document.getElementById('existing-flow')?.classList.add('hide');
  } else if(app.flow === 'trade2025'){
    document.getElementById('sec-2')?.classList.remove('hide');
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
  bindSyntheticInputs();
  bindSyntheticNav();
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
    const activeKey = app.submitted ? (app.flow === 'new' ? 's8' : 'mrv') : (app.flow === 'new' ? 's2w' : 's2w');
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
