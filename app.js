/* ── Change Logger ─────────────────────────────────────────────────────────── */
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw4r8BhwrxqTOhGZncJ6fEHfiJFcrA_hWUwgd3JqFgCFrwJAF3V4o62-O0at-MLPOVY/exec';

// Google Drive folder ID where reviewed WEGs are saved
const DRIVE_FOLDER_ID = '19pOrBUKYZKx6ZHtfXt12Qn3V60BeIhPj';

function getReviewerName() {
  let name = localStorage.getItem('weg_reviewer_name');
  if (!name) {
    name = prompt('Enter your name for change tracking (saved locally):') || 'anonymous';
    localStorage.setItem('weg_reviewer_name', name);
  }
  return name;
}

function logChange(guideId, stepId, field, oldValue, newValue) {
  if (String(oldValue) === String(newValue)) return;
  fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      timestamp : new Date().toISOString().slice(0, 19).replace('T', ' '),
      user      : getReviewerName(),
      guide_id  : String(guideId  ?? ''),
      step_id   : String(stepId   ?? ''),
      field,
      old_value : String(oldValue ?? ''),
      new_value : String(newValue ?? ''),
    }),
  }).catch(() => {});
}

/* ── State ─────────────────────────────────────────────────────────────────── */
let weg          = null;
let fileName     = 'weg.json';
let currentView  = 'header';
let currentStep  = 0;
let changeCount  = 0;

/* ── Image / Annotation State ─────────────────────────────────────────────── */
let imageFiles   = new Map();   // basename → { url: objectURL }
let stepImgCache = {};          // stepIndex → [{ name, url }]
const annot = {
  stepIndex:  null,
  partIndex:  null,   // which part is being annotated (null = none selected)
  drawing:    false,
  sx: 0, sy: 0,       // drag start in canvas pixels
};

/* ── Part color palette ────────────────────────────────────────────────────── */
const COLORS = [
  '#6366f1','#10b981','#f59e0b','#ef4444',
  '#ec4899','#14b8a6','#8b5cf6','#f97316',
  '#06b6d4','#84cc16','#fb7185','#a78bfa',
];
const pc = (i) => COLORS[i % COLORS.length];

/* ── Boot ──────────────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  setupUpload();
  setupFolderPicker();
  document.getElementById('back-btn').addEventListener('click', goHome);
  document.getElementById('download-btn').addEventListener('click', downloadWEG);
  document.getElementById('save-cloud-btn').addEventListener('click', saveToCloud);
  document.getElementById('nav-header').addEventListener('click', showHeaderView);
});

/* ════════════════════════════════════════════════════════════════════════════
   UPLOAD
   ════════════════════════════════════════════════════════════════════════════ */
function setupUpload() {
  const dz    = document.getElementById('drop-zone');
  const input = document.getElementById('file-input');
  const btn   = document.getElementById('browse-btn');

  btn.addEventListener('click', (e) => { e.stopPropagation(); input.click(); });
  dz.addEventListener('click', () => input.click());
  dz.addEventListener('dragover',  (e) => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', ()  => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
  input.addEventListener('change', () => { if (input.files[0]) handleFile(input.files[0]); });
}

function handleFile(file) {
  if (!file.name.endsWith('.json')) { showToast('Please upload a .json file', 'error'); return; }
  fileName = file.name;
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const parsed  = JSON.parse(e.target.result);
      const guideId = String(parsed.header?.guide_id ?? '');

      if (guideId) {
        const exists = await checkGuideIdExists(guideId);
        if (exists) {
          const proceed = confirm(
            `Guide ID ${guideId} already exists in the shared Drive.\n\nChoose a different WEG or click OK to open it anyway.`
          );
          if (!proceed) return;
        }
      }

      weg = parsed;
      changeCount = 0;
      imageFiles.clear();
      stepImgCache = {};
      document.getElementById('imgs-status').textContent = '';
      document.getElementById('imgs-status').className = 'imgs-status';
      openReviewScreen();
    } catch {
      showToast('Could not parse JSON', 'error');
    }
  };
  reader.readAsText(file);
}

async function checkGuideIdExists(guideId) {
  try {
    const url = `${APPS_SCRIPT_URL}?type=check_guide_id&guide_id=${encodeURIComponent(guideId)}&root_folder_id=${encodeURIComponent(DRIVE_FOLDER_ID)}`;
    const res  = await fetch(url);
    const data = await res.json();
    return data.exists === true;
  } catch {
    return false; // if check fails, don't block the user
  }
}

/* ════════════════════════════════════════════════════════════════════════════
   FOLDER PICKER  (for images)
   ════════════════════════════════════════════════════════════════════════════ */
function setupFolderPicker() {
  const btn   = document.getElementById('load-imgs-btn');
  const input = document.getElementById('imgs-folder-input');
  btn.addEventListener('click', () => input.click());
  input.addEventListener('change', () => {
    if (input.files.length) handleImgsFolder(input.files);
  });
}

function handleImgsFolder(files) {
  if (!weg) { showToast('Upload a WEG file first', 'error'); return; }

  const guideId = weg?.header?.guide_id;
  const prefix  = guideId != null ? `guide_${guideId}_step_` : null;

  // Revoke old object URLs
  imageFiles.forEach(v => URL.revokeObjectURL(v.url));
  imageFiles.clear();
  stepImgCache = {};

  let matched = 0;
  let total   = 0;

  for (const f of files) {
    if (!/\.(jpe?g|png|webp|gif)$/i.test(f.name)) continue;
    total++;
    if (prefix && !f.name.startsWith(prefix)) continue;
    imageFiles.set(f.name, { url: URL.createObjectURL(f) });
    matched++;
  }

  const status = document.getElementById('imgs-status');
  if (matched === 0 && total > 0) {
    status.textContent = `⚠ No images match guide ${guideId}`;
    status.className   = 'imgs-status warn';
    showToast(`No images matched guide_${guideId}_step_*`, 'error');
  } else if (matched === 0) {
    status.textContent = '⚠ No image files found';
    status.className   = 'imgs-status warn';
  } else {
    status.textContent = `✓ ${matched} images — guide ${guideId}`;
    status.className   = 'imgs-status ok';
    showToast(`Loaded ${matched} images for guide ${guideId}`, 'success');
  }

  // Re-render current step if visible
  if (currentView === 'step') renderStep(currentStep);
}

/* Return sorted list of images for a given step_id */
function getStepImages(stepId) {
  const guideId = weg?.header?.guide_id;
  if (!guideId) return [];
  const prefix = `guide_${guideId}_step_${stepId}_`;
  const hits = [];
  for (const [name, entry] of imageFiles) {
    if (name.startsWith(prefix)) hits.push({ name, url: entry.url });
  }
  hits.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  return hits;
}

/* ════════════════════════════════════════════════════════════════════════════
   NAVIGATION
   ════════════════════════════════════════════════════════════════════════════ */
function goHome() {
  document.getElementById('upload-screen').classList.remove('hidden');
  document.getElementById('review-screen').classList.add('hidden');
  weg = null;
}

function openReviewScreen() {
  document.getElementById('upload-screen').classList.add('hidden');
  document.getElementById('review-screen').classList.remove('hidden');
  document.getElementById('file-badge').textContent = fileName;
  document.getElementById('guide-title-display').textContent = weg?.header?.title || 'WEG';
  updateChangePill();
  renderSidebar();
  showHeaderView();
}

function showHeaderView() {
  currentView = 'header';
  document.getElementById('nav-header').classList.add('active');
  document.querySelectorAll('.step-nav-item').forEach(b => b.classList.remove('active'));
  document.getElementById('header-view').classList.remove('hidden');
  document.getElementById('step-view').classList.add('hidden');
  renderHeader();
}

function selectStep(index) {
  currentView = 'step';
  currentStep = index;
  annot.partIndex = null;
  document.getElementById('nav-header').classList.remove('active');
  document.querySelectorAll('.step-nav-item').forEach((b, i) => b.classList.toggle('active', i === index));
  document.getElementById('header-view').classList.add('hidden');
  document.getElementById('step-view').classList.remove('hidden');
  document.querySelector('.content').scrollTop = 0;
  renderStep(index);
}

/* ════════════════════════════════════════════════════════════════════════════
   SIDEBAR
   ════════════════════════════════════════════════════════════════════════════ */
function renderSidebar() {
  const nav = document.getElementById('step-nav');
  nav.innerHTML = '';
  (weg.steps || []).forEach((step, i) => {
    const btn = document.createElement('button');
    btn.className = 'step-nav-item';
    btn.innerHTML = `
      <span class="step-num">${step.step_id ?? i + 1}</span>
      <span class="step-nav-label">${esc(step.task_name || `Step ${step.step_id ?? i + 1}`)}</span>`;
    btn.addEventListener('click', () => selectStep(i));
    nav.appendChild(btn);
  });
}

function updateSidebarLabel(index) {
  const items = document.querySelectorAll('.step-nav-item');
  if (items[index]) {
    const step = weg.steps[index];
    items[index].querySelector('.step-nav-label').textContent =
      step.task_name || `Step ${step.step_id ?? index + 1}`;
  }
}

/* ════════════════════════════════════════════════════════════════════════════
   HEADER VIEW
   ════════════════════════════════════════════════════════════════════════════ */
function renderHeader() {
  const h    = weg.header || {};
  const view = document.getElementById('header-view');

  view.innerHTML = `
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Guide Metadata</span>
        ${h.guide_id != null ? `<span class="card-badge">ID ${h.guide_id}</span>` : ''}
      </div>
      <div class="card-body">
        <div class="field-group">
          <div class="field">
            <span class="field-label">Title</span>
            <input class="editable" type="text" value="${esc(h.title||'')}" data-path="header.title" />
          </div>
          <div class="field">
            <span class="field-label">Description</span>
            <textarea class="editable" rows="3" data-path="header.description">${esc(h.description||'')}</textarea>
          </div>
          <div class="field">
            <span class="field-label">Source URL</span>
            <input class="editable" type="text" value="${esc(h.source_url||'')}" data-path="header.source_url" />
          </div>
          <div class="field">
            <span class="field-label">Guide ID</span>
            <input class="editable" type="number" value="${h.guide_id??''}" data-path="header.guide_id" style="width:140px" />
          </div>
        </div>
      </div>
    </div>

    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Toolbox</span>
        <span class="card-badge">${(h.toolbox||[]).length} tools</span>
      </div>
      <div class="card-body">${makeList(h.toolbox||[], 'header.toolbox', 'Add tool…')}</div>
    </div>

    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Parts List</span>
        <span class="card-badge">${(h.parts_list||[]).length} parts</span>
      </div>
      <div class="card-body">${makeList(h.parts_list||[], 'header.parts_list', 'Add part…')}</div>
    </div>`;

  view.querySelectorAll('.editable[data-path]').forEach(el => {
    el.addEventListener('focus', () => { el.dataset.prev = el.value; });
    el.addEventListener('change', () => {
      const old = el.dataset.prev ?? '';
      let v = el.value;
      if (el.type === 'number') v = v !== '' ? Number(v) : null;
      setPath(weg, el.dataset.path, v);
      if (el.dataset.path === 'header.title')
        document.getElementById('guide-title-display').textContent = v;
      logChange(weg.header?.guide_id, 'header', el.dataset.path, old, el.value);
      bump();
    });
  });
  bindListEvents(view);
}

/* ════════════════════════════════════════════════════════════════════════════
   STEP VIEW
   ════════════════════════════════════════════════════════════════════════════ */
function renderStep(index) {
  const step = weg.steps[index];
  if (!step) return;
  const view = document.getElementById('step-view');

  view.innerHTML = `
    <!-- Step Hero -->
    <div class="step-hero">
      <div class="step-hero-top">
        <div class="step-number">${step.step_id ?? index + 1}</div>
        <div class="step-hero-title-wrap">
          <span class="field-label">Task Name</span>
          <input class="editable" type="text" value="${esc(step.task_name||'')}"
            data-sf="task_name" style="font-size:17px;font-weight:700" />
        </div>
      </div>
      <div class="field">
        <span class="field-label">Description</span>
        <textarea class="editable" rows="4" data-sf="description" style="width:100%">${esc(step.description||'')}</textarea>
      </div>
      <div class="field" style="margin-top:12px">
        <span class="field-label">Geometric Location</span>
        <input class="editable" type="text" value="${esc(step.geometric_location||'')}"
          data-sf="geometric_location" style="width:100%" />
      </div>
    </div>

    <!-- Image Annotation -->
    ${makeAnnotationCard(step, index)}

    <!-- Actions -->
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Actions</span>
        <span class="card-badge">${(step.actions||[]).length}</span>
      </div>
      <div class="card-body">${makeList(step.actions||[], `steps.${index}.actions`, 'Add action…')}</div>
    </div>

    <!-- Action Quadruples -->
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Action Quadruples</span>
        <span class="card-badge">${(step.action_quadruples||[]).length}</span>
      </div>
      <div class="card-body">${makeAQTable(step.action_quadruples||[], index)}</div>
    </div>

    <!-- Hints -->
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Hints & Notes</span>
        <span class="card-badge">${(step.hints||[]).length}</span>
      </div>
      <div class="card-body">${makeList(step.hints||[], `steps.${index}.hints`, 'Add hint…')}</div>
    </div>

    ${step.primary_part ? `
    <div class="section-card">
      <div class="card-header"><span class="card-title">Primary Part</span></div>
      <div class="card-body">
        <div class="primary-part-ref">
          <span>ID #${step.primary_part.part_id} — <strong>${esc(step.primary_part.name)}</strong></span>
          <span style="color:var(--text3);font-size:12px">(edit in Parts below)</span>
        </div>
      </div>
    </div>` : ''}

    <!-- Parts All -->
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Parts / Components</span>
        <span class="card-badge">${(step.parts_all||[]).length} detected</span>
      </div>
      <div class="card-body">${makeParts(step.parts_all||[], index)}</div>
    </div>`;

  bindStepEvents(view, index);
}

function bindStepEvents(view, index) {
  /* Simple step fields */
  view.querySelectorAll('[data-sf]').forEach(el => {
    el.addEventListener('focus', () => { el.dataset.prev = el.value; });
    el.addEventListener('change', () => {
      const old = el.dataset.prev ?? '';
      weg.steps[index][el.dataset.sf] = el.value;
      if (el.dataset.sf === 'task_name') updateSidebarLabel(index);
      logChange(weg.header?.guide_id, weg.steps[index]?.step_id, el.dataset.sf, old, el.value);
      bump();
    });
  });

  /* AQ table */
  view.querySelectorAll('.aq-input, .aq-select').forEach(el => {
    el.addEventListener('focus', () => { el.dataset.prev = el.value; });
    el.addEventListener('change', () => {
      const row = parseInt(el.dataset.row), field = el.dataset.field;
      const aq  = weg.steps[index].action_quadruples;
      if (!aq?.[row]) return;
      const old = el.dataset.prev ?? '';
      let v = el.value;
      if (field === 'hands')   v = parseInt(v);
      if (field === 'part_id') v = v !== '' ? parseInt(v) : null;
      if ((field === 'tool' || field === 'precise_action' || field === 'part') && v === '') v = null;
      aq[row][field] = v;
      logChange(weg.header?.guide_id, weg.steps[index]?.step_id, field, old, el.value);
      bump();
    });
  });

  view.querySelector('.add-row-btn')?.addEventListener('click', () => {
    if (!weg.steps[index].action_quadruples) weg.steps[index].action_quadruples = [];
    weg.steps[index].action_quadruples.push({
      action: 'new_action', precise_action: null, tool: null,
      component: 'component', part: null, part_id: null,
      hands: 1, full_action: 'New full action description.'
    });
    bump(); renderStep(index);
  });

  view.querySelectorAll('.del-row-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const row = parseInt(btn.dataset.row);
      const q   = weg.steps[index].action_quadruples?.[row];
      if (q) logChange(weg.header?.guide_id, weg.steps[index]?.step_id, 'quadruple_deleted', `[${q.action}] ${q.component}`, '');
      weg.steps[index].action_quadruples.splice(row, 1);
      bump(); renderStep(index);
    });
  });

  /* Parts */
  view.querySelectorAll('.part-input').forEach(el => {
    el.addEventListener('focus', () => { el.dataset.prev = el.value; });
    el.addEventListener('change', () => {
      const i = parseInt(el.dataset.part), field = el.dataset.field;
      const p = weg.steps[index].parts_all?.[i];
      if (!p) return;
      const old = el.dataset.prev ?? '';
      if (field === 'confidence') p[field] = el.value !== '' ? parseFloat(el.value) : null;
      else if (field === 'part_id') p[field] = parseInt(el.value) || 1;
      else p[field] = el.value;
      logChange(weg.header?.guide_id, weg.steps[index]?.step_id, `part_${field}`, old, el.value);
      bump();
    });
  });

  view.querySelectorAll('.bbox-input').forEach(el => {
    el.addEventListener('focus', () => { el.dataset.prev = el.value; });
    el.addEventListener('change', () => {
      const i = parseInt(el.dataset.part), coord = el.dataset.bbox;
      const p = weg.steps[index].parts_all?.[i];
      if (!p) return;
      if (!p.bbox) p.bbox = { x1:0, y1:0, x2:0, y2:0 };
      const old = el.dataset.prev ?? '';
      p.bbox[coord] = parseInt(el.value) || 0;
      logChange(weg.header?.guide_id, weg.steps[index]?.step_id, `bbox_${coord}:${p.name}`, old, el.value);
      bump();
      redrawCanvas(index);  // live update the canvas
    });
  });

  view.querySelectorAll('.part-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const i = parseInt(btn.dataset.part);
      const p = weg.steps[index].parts_all?.[i];
      if (p) logChange(weg.header?.guide_id, weg.steps[index]?.step_id, 'part_deleted', p.name, '');
      weg.steps[index].parts_all.splice(i, 1);
      bump(); renderStep(index);
    });
  });

  view.querySelector('.add-part-btn')?.addEventListener('click', () => {
    if (!weg.steps[index].parts_all) weg.steps[index].parts_all = [];
    const newId = weg.steps[index].parts_all.length + 1;
    weg.steps[index].parts_all.push({
      part_id: newId, name: 'new_part',
      bbox: { x1:0, y1:0, x2:100, y2:100 },
      confidence: 0.9, image_path: ''
    });
    bump(); renderStep(index);
  });

  /* Annotation */
  bindAnnotationEvents(view, index);
  bindListEvents(view);
}

/* ════════════════════════════════════════════════════════════════════════════
   ANNOTATION CARD  (HTML builder)
   ════════════════════════════════════════════════════════════════════════════ */
function makeAnnotationCard(step, stepIndex) {
  const stepId = step.step_id ?? stepIndex + 1;
  const parts  = step.parts_all || [];

  /* No folder loaded yet */
  if (imageFiles.size === 0) {
    return `
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Image Annotation</span>
      </div>
      <div class="card-body annotation-body">
        <div class="no-imgs-msg">
          <span>Load the guide's <strong>imgs</strong> folder to annotate bounding boxes visually.</span>
          <button class="add-item-btn" id="inline-load-imgs-btn" style="width:auto">Load Images Folder</button>
        </div>
      </div>
    </div>`;
  }

  const imgs = getStepImages(stepId);

  /* Folder loaded but no images for this step */
  if (imgs.length === 0) {
    return `
    <div class="section-card">
      <div class="card-header">
        <span class="card-title">Image Annotation</span>
      </div>
      <div class="card-body annotation-body">
        <p style="color:var(--text3);font-size:13px">
          No images found for step ${stepId}
          <span style="color:var(--text3);font-size:11px">(expected: guide_${weg?.header?.guide_id}_step_${stepId}_1.jpg …)</span>
        </p>
      </div>
    </div>`;
  }

  /* Cache the image list for this step (used by canvas navigation) */
  stepImgCache[stepIndex] = imgs;

  const legend = parts.length
    ? parts.map((p, i) => `
        <button class="part-legend-item" data-part-annot="${i}" style="--pc:${pc(i)}">
          <span class="legend-dot" style="background:${pc(i)}"></span>
          <span class="legend-name">${esc(p.name || `Part ${p.part_id ?? i+1}`)}</span>
        </button>`).join('')
    : `<p style="color:var(--text3);font-size:12px;margin:0">No parts yet — click + Add Part</p>`;

  const multiNav = imgs.length > 1 ? `
    <div class="img-nav">
      <button class="img-nav-btn" id="img-prev-${stepIndex}">←</button>
      <span class="img-nav-label" id="img-nav-lbl-${stepIndex}">1 / ${imgs.length}</span>
      <button class="img-nav-btn" id="img-next-${stepIndex}">→</button>
    </div>` : '';

  return `
  <div class="section-card annotation-card">
    <div class="card-header">
      <span class="card-icon">🖼️</span>
      <span class="card-title">Image Annotation</span>
      ${multiNav}
      <span class="card-badge" style="${imgs.length > 1 ? '' : 'margin-left:auto'}">${imgs.length} image${imgs.length > 1 ? 's' : ''}</span>
    </div>
    <div class="card-body annotation-body">
      <div class="annot-layout">
        <div class="canvas-outer">
          <div class="canvas-container" id="canvas-wrap-${stepIndex}">
            <img id="annot-img-${stepIndex}"
              src="${imgs[0].url}"
              class="annot-img"
              draggable="false" />
            <canvas id="annot-canvas-${stepIndex}" class="annot-canvas"></canvas>
          </div>
          <div class="annot-hint" id="annot-hint-${stepIndex}">
            Select a part on the right, then drag on the image to draw its bounding box
          </div>
        </div>
        <div class="part-legend-panel" id="part-legend-${stepIndex}">
          <div class="legend-title">Parts</div>
          ${legend}
          <div class="add-part-row">
            <input class="add-part-name-input" type="text" placeholder="Part name…" />
            <button class="add-part-legend-btn">Add</button>
          </div>
          <div class="legend-hint">Click to select · draw bbox · click again to deselect</div>
        </div>
      </div>
    </div>
  </div>`;
}

/* ════════════════════════════════════════════════════════════════════════════
   ANNOTATION EVENTS + CANVAS LOGIC
   ════════════════════════════════════════════════════════════════════════════ */
function bindAnnotationEvents(view, stepIndex) {
  /* Inline "load folder" button (shown when no images loaded) */
  view.querySelector('#inline-load-imgs-btn')?.addEventListener('click', () => {
    document.getElementById('imgs-folder-input').click();
  });

  /* Image navigation */
  view.querySelector(`#img-prev-${stepIndex}`)?.addEventListener('click', () => shiftAnnotImg(stepIndex, -1));
  view.querySelector(`#img-next-${stepIndex}`)?.addEventListener('click', () => shiftAnnotImg(stepIndex, +1));

  /* Part legend selection */
  view.querySelectorAll('.part-legend-item').forEach((btn, i) => {
    btn.addEventListener('click', () => {
      if (annot.partIndex === i && annot.stepIndex === stepIndex) {
        deselectAnnotPart(stepIndex);
      } else {
        selectAnnotPart(stepIndex, i);
      }
    });
  });

  /* Add Part from legend panel */
  view.querySelector('.add-part-legend-btn')?.addEventListener('click', () => {
    const input = view.querySelector('.add-part-name-input');
    const name  = input?.value.trim();
    if (!name) { input?.focus(); return; }
    if (!weg.steps[stepIndex].parts_all) weg.steps[stepIndex].parts_all = [];
    const newId = weg.steps[stepIndex].parts_all.length + 1;
    weg.steps[stepIndex].parts_all.push({
      part_id: newId, name,
      bbox: { x1: 0, y1: 0, x2: 100, y2: 100 },
      confidence: 0.9, image_path: '',
    });
    bump(); renderStep(stepIndex);
  });

  view.querySelector('.add-part-name-input')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') view.querySelector('.add-part-legend-btn')?.click();
  });

  /* Init canvas after image loads */
  const img = document.getElementById(`annot-img-${stepIndex}`);
  if (img) {
    if (img.complete && img.naturalWidth) {
      initCanvas(stepIndex);
    } else {
      img.addEventListener('load', () => initCanvas(stepIndex), { once: true });
    }
  }
}

function initCanvas(stepIndex) {
  const img    = document.getElementById(`annot-img-${stepIndex}`);
  const canvas = document.getElementById(`annot-canvas-${stepIndex}`);
  if (!img || !canvas) return;

  /* Sync canvas pixel dimensions to displayed image size */
  const syncSize = () => {
    canvas.width  = img.offsetWidth;
    canvas.height = img.offsetHeight;
    redrawCanvas(stepIndex);
  };
  syncSize();

  /* Redraw on window resize */
  const ro = new ResizeObserver(syncSize);
  ro.observe(img);

  /* ── Mouse drawing ── */
  canvas.addEventListener('mousedown', (e) => {
    if (annot.partIndex === null || annot.stepIndex !== stepIndex) return;
    annot.drawing = true;
    const r = canvas.getBoundingClientRect();
    annot.sx = e.clientX - r.left;
    annot.sy = e.clientY - r.top;
  });

  canvas.addEventListener('mousemove', (e) => {
    if (!annot.drawing) return;
    const r  = canvas.getBoundingClientRect();
    const cx = e.clientX - r.left;
    const cy = e.clientY - r.top;
    redrawCanvas(stepIndex);
    /* Preview rect */
    const ctx = canvas.getContext('2d');
    const col = pc(annot.partIndex);
    ctx.save();
    ctx.strokeStyle = col;
    ctx.fillStyle   = col + '22';
    ctx.lineWidth   = 2;
    ctx.setLineDash([5, 4]);
    const rx = Math.min(annot.sx, cx), ry = Math.min(annot.sy, cy);
    const rw = Math.abs(cx - annot.sx), rh = Math.abs(cy - annot.sy);
    ctx.fillRect(rx, ry, rw, rh);
    ctx.strokeRect(rx, ry, rw, rh);
    ctx.restore();
  });

  canvas.addEventListener('mouseup', (e) => {
    if (!annot.drawing) return;
    annot.drawing = false;
    const r  = canvas.getBoundingClientRect();
    const ex = e.clientX - r.left;
    const ey = e.clientY - r.top;

    const w = Math.abs(ex - annot.sx);
    const h = Math.abs(ey - annot.sy);
    if (w < 4 || h < 4) return;   // too small — ignore accidental clicks

    /* Convert display coords → original image pixel coords */
    const scaleX = img.naturalWidth  / canvas.width;
    const scaleY = img.naturalHeight / canvas.height;
    const bbox = {
      x1: Math.round(Math.min(annot.sx, ex) * scaleX),
      y1: Math.round(Math.min(annot.sy, ey) * scaleY),
      x2: Math.round(Math.max(annot.sx, ex) * scaleX),
      y2: Math.round(Math.max(annot.sy, ey) * scaleY),
    };

    /* Save to WEG */
    const part = weg.steps[stepIndex]?.parts_all?.[annot.partIndex];
    if (part) {
      const oldBbox = JSON.stringify(part.bbox ?? {});
      part.bbox = bbox;
      logChange(weg.header?.guide_id, weg.steps[stepIndex]?.step_id, `bbox:${part.name}`, oldBbox, JSON.stringify(bbox));
      bump();
      syncBboxInputs(stepIndex, annot.partIndex, bbox);
      showToast(`✓ BBox set for "${part.name}"`, 'success');
    }
    redrawCanvas(stepIndex);
  });

  /* Cancel drawing on mouse leave */
  canvas.addEventListener('mouseleave', () => {
    if (annot.drawing) { annot.drawing = false; redrawCanvas(stepIndex); }
  });
}

/* Draw all bboxes on the canvas */
function redrawCanvas(stepIndex) {
  const img    = document.getElementById(`annot-img-${stepIndex}`);
  const canvas = document.getElementById(`annot-canvas-${stepIndex}`);
  if (!img || !canvas || !img.naturalWidth) return;

  const ctx    = canvas.getContext('2d');
  const parts  = weg.steps[stepIndex]?.parts_all || [];
  const scaleX = canvas.width  / img.naturalWidth;
  const scaleY = canvas.height / img.naturalHeight;

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  parts.forEach((p, i) => {
    const bbox = p.bbox;
    if (!bbox || (bbox.x2 <= bbox.x1) || (bbox.y2 <= bbox.y1)) return;

    const x = bbox.x1 * scaleX;
    const y = bbox.y1 * scaleY;
    const w = (bbox.x2 - bbox.x1) * scaleX;
    const h = (bbox.y2 - bbox.y1) * scaleY;
    const col     = pc(i);
    const isActive = annot.partIndex === i && annot.stepIndex === stepIndex;

    ctx.save();
    ctx.strokeStyle = col;
    ctx.fillStyle   = col + (isActive ? '33' : '18');
    ctx.lineWidth   = isActive ? 3 : 1.5;
    ctx.setLineDash(isActive ? [] : [4, 3]);
    ctx.fillRect(x, y, w, h);
    ctx.strokeRect(x, y, w, h);
    ctx.restore();

    /* Label */
    const label = p.name || `Part ${p.part_id ?? i+1}`;
    ctx.save();
    ctx.font = `${isActive ? 'bold ' : ''}${isActive ? 12 : 11}px Inter, sans-serif`;
    const tw = ctx.measureText(label).width;
    const lx = x, ly = y > 20 ? y - 4 : y + h + 14;
    ctx.fillStyle = 'rgba(0,0,0,0.7)';
    ctx.fillRect(lx, ly - 13, tw + 8, 16);
    ctx.fillStyle = col;
    ctx.fillText(label, lx + 4, ly);
    ctx.restore();
  });
}

function selectAnnotPart(stepIndex, partIndex) {
  annot.stepIndex  = stepIndex;
  annot.partIndex  = partIndex;

  document.querySelectorAll('.part-legend-item').forEach((btn, i) =>
    btn.classList.toggle('active', i === partIndex));

  const canvas = document.getElementById(`annot-canvas-${stepIndex}`);
  if (canvas) canvas.classList.add('drawing');

  const hint = document.getElementById(`annot-hint-${stepIndex}`);
  const part = weg.steps[stepIndex]?.parts_all?.[partIndex];
  if (hint && part) {
    hint.textContent = `Drawing bbox for "${part.name}" — click and drag on the image`;
    hint.style.color  = pc(partIndex);
  }
  redrawCanvas(stepIndex);
}

function deselectAnnotPart(stepIndex) {
  annot.partIndex = null;
  annot.drawing   = false;

  document.querySelectorAll('.part-legend-item').forEach(b => b.classList.remove('active'));

  const canvas = document.getElementById(`annot-canvas-${stepIndex}`);
  if (canvas) canvas.classList.remove('drawing');

  const hint = document.getElementById(`annot-hint-${stepIndex}`);
  if (hint) {
    hint.textContent = 'Select a part on the right, then drag on the image to draw its bounding box';
    hint.style.color  = '';
  }
  redrawCanvas(stepIndex);
}

/* After drawing, sync bbox inputs without re-rendering the whole step */
function syncBboxInputs(stepIndex, partIndex, bbox) {
  ['x1','y1','x2','y2'].forEach(k => {
    const el = document.querySelector(`.bbox-input[data-part="${partIndex}"][data-bbox="${k}"]`);
    if (el) el.value = bbox[k];
  });
}

/* Switch image in multi-image steps */
function shiftAnnotImg(stepIndex, delta) {
  const imgs  = stepImgCache[stepIndex];
  const img   = document.getElementById(`annot-img-${stepIndex}`);
  const label = document.getElementById(`img-nav-lbl-${stepIndex}`);
  if (!imgs || !img) return;

  const cur = parseInt(img.dataset.imgIdx || '0');
  const nxt = (cur + delta + imgs.length) % imgs.length;
  img.dataset.imgIdx = nxt;
  img.onload = () => {
    const canvas = document.getElementById(`annot-canvas-${stepIndex}`);
    if (canvas) { canvas.width = img.offsetWidth; canvas.height = img.offsetHeight; }
    redrawCanvas(stepIndex);
  };
  img.src = imgs[nxt].url;
  if (label) label.textContent = `${nxt + 1} / ${imgs.length}`;
}

/* ════════════════════════════════════════════════════════════════════════════
   HTML BUILDERS
   ════════════════════════════════════════════════════════════════════════════ */
function makeList(items, path, placeholder) {
  const rows = items.map((item, i) => `
    <div class="list-item">
      <textarea class="list-item-text"
        data-list-path="${escAttr(path)}" data-list-index="${i}"
        oninput="autoResize(this)" rows="1">${esc(item)}</textarea>
      <button class="list-item-del"
        data-list-path="${escAttr(path)}" data-list-index="${i}">×</button>
    </div>`).join('');

  return `<div class="list-field" data-list-path="${escAttr(path)}">
    ${rows}
    <button class="add-item-btn" data-add-path="${escAttr(path)}">＋ ${placeholder}</button>
  </div>`;
}

function makeAQTable(quads, stepIndex) {
  if (!quads.length) {
    return `<p style="color:var(--text3);font-size:13px;margin-bottom:10px">No quadruples yet.</p>
            <button class="add-row-btn">＋ Add Quadruple</button>`;
  }

  const heads = ['Action','Precise Action','Tool','Component','Part','Part ID','Hands','Full Action',''];
  const rows  = quads.map((q, i) => `
    <tr class="aq-row">
      <td class="aq-cell"><input class="aq-input" type="text" value="${esc(q.action||'')}" data-row="${i}" data-field="action" /></td>
      <td class="aq-cell"><input class="aq-input" type="text" value="${esc(q.precise_action||'')}" data-row="${i}" data-field="precise_action" placeholder="—" /></td>
      <td class="aq-cell"><input class="aq-input" type="text" value="${esc(q.tool||'')}" data-row="${i}" data-field="tool" placeholder="none" /></td>
      <td class="aq-cell"><input class="aq-input" type="text" value="${esc(q.component||'')}" data-row="${i}" data-field="component" /></td>
      <td class="aq-cell"><input class="aq-input" type="text" value="${esc(q.part||'')}" data-row="${i}" data-field="part" placeholder="—" /></td>
      <td class="aq-cell"><input class="aq-input" type="number" value="${q.part_id??''}" data-row="${i}" data-field="part_id" style="width:50px" placeholder="—" /></td>
      <td class="aq-cell">
        <select class="aq-select" data-row="${i}" data-field="hands">
          <option value="0" ${q.hands===0?'selected':''}>0 🤲</option>
          <option value="1" ${q.hands===1?'selected':''}>1 ✋</option>
          <option value="2" ${q.hands===2?'selected':''}>2 👐</option>
        </select>
      </td>
      <td class="aq-cell" style="min-width:200px;max-width:300px">
        <input class="aq-input" type="text" value="${esc(q.full_action||'')}" data-row="${i}" data-field="full_action" />
      </td>
      <td><button class="del-row-btn" data-row="${i}">×</button></td>
    </tr>`).join('');

  return `<div class="aq-scroll">
    <table class="aq-table">
      <thead><tr>${heads.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>
  <button class="add-row-btn">＋ Add Quadruple</button>`;
}

function makeParts(parts, stepIndex) {
  if (!parts.length) {
    return `<p style="color:var(--text3);font-size:13px;margin-bottom:10px">No parts detected.</p>
            <button class="add-part-btn add-item-btn">＋ Add Part</button>`;
  }

  const cards = parts.map((p, i) => {
    const bb   = p.bbox || {};
    const conf = p.confidence ?? 0;
    return `
    <div class="part-card" style="border-left: 3px solid ${pc(i)}">
      <div class="part-card-top">
        <span class="part-id-badge" style="background:${pc(i)}">Part #${p.part_id ?? i+1}</span>
        <button class="part-del-btn" data-part="${i}">×</button>
      </div>
      <div class="part-field">
        <div class="part-label">Name</div>
        <input class="part-input" type="text" value="${esc(p.name||'')}" data-part="${i}" data-field="name" />
      </div>
      <div class="part-field">
        <div class="part-label">Confidence</div>
        <div class="conf-bar-wrap">
          <div class="conf-bar"><div class="conf-bar-fill" style="width:${Math.round(conf*100)}%"></div></div>
          <input class="part-input" type="number" min="0" max="1" step="0.01"
            value="${conf}" data-part="${i}" data-field="confidence" style="width:70px;text-align:right" />
        </div>
      </div>
      <div class="part-field">
        <div class="part-label">Bounding Box <span style="color:var(--text3);font-size:9px">(or draw above)</span></div>
        <div class="bbox-row">
          ${['x1','y1','x2','y2'].map(k => `
          <div class="bbox-wrap">
            <span class="bbox-lbl">${k}</span>
            <input class="bbox-input" type="number" value="${bb[k]??0}" data-part="${i}" data-bbox="${k}" />
          </div>`).join('')}
        </div>
      </div>
      <div class="part-field">
        <div class="part-label">Image Path</div>
        <input class="part-input" type="text" value="${esc(p.image_path||'')}"
          data-part="${i}" data-field="image_path" style="font-size:11px" />
      </div>
    </div>`;
  }).join('');

  return `<div class="parts-grid">${cards}</div>
          <button class="add-part-btn add-item-btn" style="margin-top:14px">＋ Add Part</button>`;
}

/* ════════════════════════════════════════════════════════════════════════════
   LIST EVENT BINDING
   ════════════════════════════════════════════════════════════════════════════ */
function bindListEvents(container) {
  container.querySelectorAll('.list-item-text').forEach(el => {
    autoResize(el);
    el.addEventListener('change', () => {
      const arr = getPath(weg, el.dataset.listPath);
      if (arr) { arr[parseInt(el.dataset.listIndex)] = el.value; bump(); }
    });
  });

  container.querySelectorAll('.list-item-del').forEach(btn => {
    btn.addEventListener('click', () => {
      const path = btn.dataset.listPath;
      const arr  = getPath(weg, path);
      if (arr) { arr.splice(parseInt(btn.dataset.listIndex), 1); bump(); rerenderSection(path); }
    });
  });

  container.querySelectorAll('[data-add-path]').forEach(btn => {
    btn.addEventListener('click', () => {
      const path = btn.dataset.addPath;
      let arr = getPath(weg, path);
      if (!arr) { setPath(weg, path, []); arr = getPath(weg, path); }
      arr.push(''); bump(); rerenderSection(path);
    });
  });
}

function rerenderSection(path) {
  if (path.startsWith('header.')) { renderHeader(); return; }
  const m = path.match(/^steps\.(\d+)\./);
  if (m) renderStep(parseInt(m[1]));
}

/* ════════════════════════════════════════════════════════════════════════════
   DOWNLOAD
   ════════════════════════════════════════════════════════════════════════════ */
function downloadWEG() {
  const json = JSON.stringify(weg, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: fileName.replace(/\.json$/i, '_reviewed.json') });
  a.click();
  URL.revokeObjectURL(url);
  showToast('✓ Downloaded!', 'success');
}

/* ════════════════════════════════════════════════════════════════════════════
   SAVE TO CLOUD (Google Drive via Apps Script)
   ════════════════════════════════════════════════════════════════════════════ */
async function saveToCloud() {
  if (!weg) return;

  const btn = document.getElementById('save-cloud-btn');
  btn.disabled = true;
  btn.textContent = 'Preparing…';

  try {
    /* ── Collect annotated images for every step ── */
    const images = [];
    for (let idx = 0; idx < weg.steps.length; idx++) {
      const step    = weg.steps[idx];
      const stepId  = step.step_id ?? idx + 1;
      const imgList = stepImgCache[idx] || getStepImages(stepId);
      if (!imgList.length) continue;

      for (let imgN = 0; imgN < imgList.length; imgN++) {
        const imgEl = new Image();
        await new Promise(resolve => {
          imgEl.onload = imgEl.onerror = resolve;
          imgEl.src = imgList[imgN].url;
        });
        if (!imgEl.naturalWidth) continue;

        /* Draw image + bbox annotations on off-screen canvas */
        const offCanvas  = document.createElement('canvas');
        offCanvas.width  = imgEl.naturalWidth;
        offCanvas.height = imgEl.naturalHeight;
        const ctx = offCanvas.getContext('2d');
        ctx.drawImage(imgEl, 0, 0);

        (step.parts_all || []).forEach((p, i) => {
          if (!p.bbox) return;
          const { x1, y1, x2, y2 } = p.bbox;
          const col = pc(i);
          ctx.save();
          ctx.strokeStyle = col;
          ctx.fillStyle   = col + '22';
          ctx.lineWidth   = 2;
          ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
          ctx.fillRect(x1,   y1, x2 - x1, y2 - y1);
          const label = p.name || `Part ${p.part_id ?? i + 1}`;
          ctx.font      = 'bold 13px sans-serif';
          const tw      = ctx.measureText(label).width;
          ctx.fillStyle = '#000a';
          ctx.fillRect(x1, y1 - 18, tw + 10, 18);
          ctx.fillStyle = col;
          ctx.fillText(label, x1 + 5, y1 - 4);
          ctx.restore();
        });

        const b64      = offCanvas.toDataURL('image/jpeg', 0.85).split(',')[1];
        const taskSlug = (step.task_name || `step_${stepId}`)
          .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 40);
        const suffix   = imgList.length > 1 ? `_${imgN + 1}` : '';
        images.push({ name: `step_${stepId}_${taskSlug}${suffix}.jpg`, data: b64 });
      }
    }

    /* ── Folder / file names ── */
    const wegName = (weg.header?.title || fileName.replace(/\.json$/i, ''))
      .replace(/[^a-zA-Z0-9_\- ]/g, '').trim().replace(/\s+/g, '_').slice(0, 60);
    const guideId = String(weg.header?.guide_id ?? 'unknown');

    btn.textContent = 'Uploading…';

    const payload = {
      type           : 'save_weg_full',
      root_folder_id : DRIVE_FOLDER_ID,
      weg_name       : wegName,
      guide_id       : guideId,
      reviewer       : getReviewerName(),
      saved_at       : new Date().toISOString().slice(0, 19).replace('T', ' '),
      weg_json       : JSON.stringify(weg, null, 2),
      images,
    };
    console.log('[saveToCloud] sending payload, images:', images.length, 'payload size (kb):', Math.round(JSON.stringify(payload).length / 1024));

    const res  = await fetch(APPS_SCRIPT_URL, {
      method : 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body   : JSON.stringify(payload),
    });
    const text = await res.text();
    console.log('[saveToCloud] response:', text);

    showToast('✓ Saved to Drive!', 'success');
    btn.textContent = '✓ Saved to Cloud';
  } catch (err) {
    showToast('Failed to save to cloud', 'error');
    btn.textContent = 'Save to Cloud';
    console.error(err);
  } finally {
    btn.disabled = false;
  }
}

/* ════════════════════════════════════════════════════════════════════════════
   HELPERS
   ════════════════════════════════════════════════════════════════════════════ */
function bump() {
  changeCount++;
  updateChangePill();
}

function updateChangePill() {
  const p = document.getElementById('change-count');
  if (p) p.textContent = `${changeCount} change${changeCount === 1 ? '' : 's'}`;
}

let _toastTimer;
function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className   = `show ${type}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => (t.className = ''), 2000);
}

function autoResize(el) {
  el.style.height = 'auto';
  el.style.height = el.scrollHeight + 'px';
}

function getPath(obj, path) {
  return path.split('.').reduce((o, k) => (o != null ? o[k] : null), obj);
}

function setPath(obj, path, value) {
  const keys = path.split('.');
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) {
    if (cur[keys[i]] == null) cur[keys[i]] = {};
    cur = cur[keys[i]];
  }
  cur[keys[keys.length - 1]] = value;
}

function esc(str) {
  return String(str ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function escAttr(str) { return String(str ?? '').replace(/"/g,'&quot;'); }
