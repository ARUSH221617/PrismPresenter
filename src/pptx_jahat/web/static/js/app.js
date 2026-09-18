// PrismPresenter Web SPA Client Application Logic - Shadcn UI & Google Gemini Style Chat

let activeTab = 'generator';
let currentGeneratedPptx = null;

// Generator State
let genSlides = [];
let genSlideIdx = 0;
let visualSlides = [];
let visualSlideIdx = 0;
let aiTestImages = [];
let aiTestIdx = 0;
let genSourceFiles = [];

// Storyboard Structure State
let structuresList = [];
let currentStructure = null;
let structSampleFiles = [];

// Template Analyzer State
let templatesList = [];
let selectedTemplateName = null;
let tplSlides = [];
let tplSlideIdx = 0;

// Manager State
let mgrGeneratedList = [];
let mgrReferenceList = [];
let mgrSelectedFile = null;
let mgrSlides = [];
let mgrSlideIdx = 0;

// Lightbox state
let lightboxSlides = [];
let lightboxIdx = 0;

// DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  initThemeMode();
  if (window.lucide) {
    lucide.createIcons();
  }
  setupDragAndDrop();
  setupKeyboardHotkeys();
  loadInitialData();
});

// -------------------------------------------------------------
// DARK / LIGHT THEME MODE CONTROLLER
// -------------------------------------------------------------
function initThemeMode() {
  const savedTheme = localStorage.getItem('prism_theme') || 'dark';
  applyThemeMode(savedTheme);
}

function applyThemeMode(theme) {
  const html = document.documentElement;
  const sunIcon = document.getElementById('theme-icon-sun');
  const moonIcon = document.getElementById('theme-icon-moon');

  if (theme === 'dark') {
    html.classList.add('dark');
    if (sunIcon) sunIcon.classList.remove('hidden');
    if (moonIcon) moonIcon.classList.add('hidden');
  } else {
    html.classList.remove('dark');
    if (sunIcon) sunIcon.classList.add('hidden');
    if (moonIcon) moonIcon.classList.remove('hidden');
  }
  localStorage.setItem('prism_theme', theme);
  refreshIcons();
}

function toggleThemeMode() {
  const currentIsDark = document.documentElement.classList.contains('dark');
  const nextTheme = currentIsDark ? 'light' : 'dark';
  applyThemeMode(nextTheme);
  showToast(`Switched to ${nextTheme === 'dark' ? 'Dark' : 'Light'} theme`, 'info', 2000);
}

function refreshIcons() {
  if (window.lucide) {
    setTimeout(() => lucide.createIcons(), 50);
  }
}

// -------------------------------------------------------------
// TOAST NOTIFICATION SYSTEM (Shadcn Toast)
// -------------------------------------------------------------
function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = 'shadcn-toast p-3 flex items-center gap-2.5 text-xs pointer-events-auto';

  let iconHtml = '<i data-lucide="info" class="w-4 h-4 text-primary"></i>';
  if (type === 'success') {
    iconHtml = '<i data-lucide="check-circle-2" class="w-4 h-4 text-emerald-400"></i>';
  } else if (type === 'error') {
    iconHtml = '<i data-lucide="alert-triangle" class="w-4 h-4 text-destructive"></i>';
  } else if (type === 'warning') {
    iconHtml = '<i data-lucide="alert-circle" class="w-4 h-4 text-amber-400"></i>';
  }

  toast.innerHTML = `
    <div class="flex-shrink-0">${iconHtml}</div>
    <div class="flex-1 font-medium text-foreground">${message}</div>
    <button onclick="this.parentElement.remove()" class="text-muted-foreground hover:text-foreground ml-2">
      <i data-lucide="x" class="w-3.5 h-3.5"></i>
    </button>
  `;

  container.appendChild(toast);
  refreshIcons();

  setTimeout(() => {
    toast.classList.add('toast-exit');
    setTimeout(() => toast.remove(), 200);
  }, duration);
}

// -------------------------------------------------------------
// DRAG & DROP FOR DOCX
// -------------------------------------------------------------
function setupDragAndDrop() {
  const dropzone = document.getElementById('docx-dropzone');
  if (!dropzone) return;

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dropzone-active');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dropzone-active');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length > 0) {
      uploadMultiFiles(Array.from(files));
    }
  }, false);

  const samplesDropzone = document.getElementById('samples-dropzone');
  if (samplesDropzone) {
    ['dragenter', 'dragover'].forEach(eventName => {
      samplesDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        samplesDropzone.classList.add('dropzone-active');
      }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
      samplesDropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        samplesDropzone.classList.remove('dropzone-active');
      }, false);
    });

    samplesDropzone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files && files.length > 0) {
        uploadStructSampleFiles(Array.from(files));
      }
    }, false);
  }
}

// -------------------------------------------------------------
// KEYBOARD HOTKEYS SYSTEM
// -------------------------------------------------------------
function setupKeyboardHotkeys() {
  document.addEventListener('keydown', (e) => {
    if (['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName)) {
      if (e.key === 'Enter' && e.ctrlKey && activeTab === 'generator') {
        startPresentationGeneration();
      }
      return;
    }

    if (e.key === 'Escape') {
      closeSlideLightbox();
      const shortcuts = document.getElementById('shortcuts-modal');
      if (shortcuts && !shortcuts.classList.contains('hidden')) {
        toggleShortcutsModal();
      }
      return;
    }

    const lightbox = document.getElementById('slide-lightbox');
    if (lightbox && !lightbox.classList.contains('hidden')) {
      if (e.key === 'ArrowLeft') navLightboxSlide(-1);
      if (e.key === 'ArrowRight') navLightboxSlide(1);
      return;
    }

    if (activeTab === 'generator') {
      if (e.key === 'ArrowLeft') navGenSlide(-1);
      if (e.key === 'ArrowRight') navGenSlide(1);
      if (e.key === 'f' || e.key === 'F') openCurrentSlideInLightbox();
    } else if (activeTab === 'manager') {
      if (e.key === 'ArrowLeft') navMgrSlide(-1);
      if (e.key === 'ArrowRight') navMgrSlide(1);
    }

    if (e.key === '1') switchTab('generator');
    if (e.key === '2') switchTab('templates');
    if (e.key === '3') switchTab('manager');
    if (e.key === '4') switchTab('components');
    if (e.key === '5') switchTab('agent');
    if (e.key === '6') switchTab('settings');
    if (e.key === '7') switchTab('help');
  });
}

function toggleShortcutsModal() {
  const modal = document.getElementById('shortcuts-modal');
  if (!modal) return;
  modal.classList.toggle('hidden');
  refreshIcons();
}

// -------------------------------------------------------------
// TAB SWITCHING (Shadcn Tabs Trigger)
// -------------------------------------------------------------
function switchTab(tabId) {
  activeTab = tabId;
  document.querySelectorAll('.shadcn-tab-trigger').forEach(btn => {
    btn.classList.remove('active');
  });

  const activeBtn = document.getElementById(`tab-btn-${tabId}`);
  if (activeBtn) {
    activeBtn.classList.add('active');
  }

  document.querySelectorAll('.tab-content').forEach(content => {
    content.classList.add('hidden');
  });

  const activeContent = document.getElementById(`tab-content-${tabId}`);
  if (activeContent) {
    activeContent.classList.remove('hidden');
  }

  // Tab Loaders & Media Management
  const brandVid = document.getElementById('brand-reveal-video');
  if (tabId === 'help') {
    if (brandVid) {
      brandVid.currentTime = 0;
      brandVid.play().catch(() => {});
    }
  } else {
    if (brandVid && !brandVid.paused) {
      brandVid.pause();
    }
  }

  if (tabId === 'templates') {
    loadTemplatesList();
    loadNoteMd();
    loadStructuresList();
  } else if (tabId === 'manager') {
    loadManagerDecks();
  } else if (tabId === 'components') {
    loadComponentsCatalog();
  } else if (tabId === 'settings') {
    loadConfigSettings();
  }

  refreshIcons();
}

function switchGenSubTab(subTabId) {
  document.querySelectorAll('.shadcn-tabs-list .shadcn-tab-trigger').forEach(btn => {
    if (btn.id.startsWith('subtab-btn-')) {
      btn.classList.remove('active');
    }
  });

  const activeBtn = document.getElementById(`subtab-btn-${subTabId}`);
  if (activeBtn) {
    activeBtn.classList.add('active');
  }

  document.querySelectorAll('.gen-subtab-content').forEach(content => {
    content.classList.add('hidden');
  });

  const activeContent = document.getElementById(`gen-subtab-${subTabId}`);
  if (activeContent) {
    activeContent.classList.remove('hidden');
  }

  refreshIcons();
}

function setSystemStatus(text, isBusy = false) {
  const badge = document.getElementById('system-status-badge');
  const txt = document.getElementById('system-status-text');
  const footer = document.getElementById('footer-status');

  if (txt) txt.innerText = text;
  if (footer) footer.innerText = text;

  if (isBusy) {
    badge.className = 'shadcn-badge shadcn-badge-default gap-1.5 py-1 animate-pulse';
  } else {
    badge.className = 'shadcn-badge shadcn-badge-secondary gap-1.5 py-1';
  }
}

// -------------------------------------------------------------
// INITIAL DATA LOADER
// -------------------------------------------------------------
async function loadInitialData() {
  try {
    loadGeneratorTemplates();
    loadGeneratorStructures();
    loadConfigBadge();
  } catch (err) {
    console.error('Error loading initial data', err);
  }
}

async function loadConfigBadge() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    if (data.success) {
      applyConfigAndMetadata(data);
    }
  } catch (e) {
    console.error(e);
  }
}

function applyConfigAndMetadata(data) {
  if (!data || !data.success) return;
  const cfg = data.config || {};
  const meta = data.model_metadata || {};

  const modelName = cfg.NINEROUTER_CHAT_MODEL ? cfg.NINEROUTER_CHAT_MODEL.split('/').pop() : 'Default';
  const badge = document.getElementById('model-badge');
  if (badge) badge.innerText = `MODEL: ${modelName}`;

  const geminiModel = document.getElementById('gemini-model-name');
  if (geminiModel) geminiModel.innerText = modelName;

  const timeoutSec = cfg.LLM_TIMEOUT || 300;
  const genTimeout = document.getElementById('gen-timeout-input');
  if (genTimeout && !genTimeout.dataset.userEdited) {
    genTimeout.value = timeoutSec;
  }

  const cfgTimeout = document.getElementById('cfg-timeout');
  if (cfgTimeout) cfgTimeout.value = timeoutSec;

  // Render detected limits
  const maxTokens = meta.max_tokens ? Number(meta.max_tokens).toLocaleString() : '--';
  const contextLength = meta.context_length ? Number(meta.context_length).toLocaleString() : '--';
  const shortMax = meta.max_tokens ? `${Math.round(meta.max_tokens / 1024)}k` : '65k';
  const shortContext = meta.context_length ? (meta.context_length >= 1000000 ? `${(meta.context_length / 1000000).toFixed(1).replace('.0','')}M` : `${Math.round(meta.context_length / 1024)}k`) : '1M';

  const genBadge = document.getElementById('gen-model-limits-badge');
  if (genBadge) {
    genBadge.innerText = `Max: ${shortMax} • Context: ${shortContext}`;
    genBadge.title = `Model: ${meta.model_id || modelName}\nMax output: ${maxTokens} tokens\nContext window: ${contextLength} tokens`;
  }

  const elMax = document.getElementById('cfg-meta-max-tokens');
  if (elMax) elMax.innerText = `${maxTokens} tokens`;

  const elCtx = document.getElementById('cfg-meta-context');
  if (elCtx) elCtx.innerText = `${contextLength} tokens`;

  const elCaps = document.getElementById('cfg-meta-caps');
  if (elCaps && meta.capabilities) {
    elCaps.innerHTML = '';
    const capList = [
      { key: 'vision', label: 'Vision OCR' },
      { key: 'tools', label: 'Tool Calling' },
      { key: 'reasoning', label: 'Extended Reasoning' },
      { key: 'search', label: 'Web Search' }
    ];
    capList.forEach(c => {
      const active = Boolean(meta.capabilities[c.key]);
      const span = document.createElement('span');
      span.className = `shadcn-badge shadcn-badge-outline text-[10px] py-0.5 px-1.5 font-mono ${active ? 'text-emerald-400 border-emerald-800/80 bg-emerald-950/30' : 'text-muted-foreground border-border'}`;
      span.innerText = `${active ? '✓' : '○'} ${c.label}`;
      elCaps.appendChild(span);
    });
  }
}

async function refreshModelMetadataUI() {
  showToast('Querying 9Router model capabilities...', 'info', 1500);
  await loadConfigBadge();
  await loadConfigSettings();
  showToast('Updated model intelligence from 9Router', 'success', 2000);
}

// -------------------------------------------------------------
// 1. SLIDE GENERATOR
// -------------------------------------------------------------
async function loadGeneratorTemplates() {
  try {
    const res = await fetch('/api/generator/templates');
    const data = await res.json();
    const select = document.getElementById('gen-template-select');
    select.innerHTML = '<option value="">✨ All Templates (Global AI Intelligent Matching)</option>';

    if (data.templates && data.templates.length > 0) {
      data.templates.forEach(tpl => {
        const opt = document.createElement('option');
        opt.value = tpl;
        opt.innerText = tpl;
        select.appendChild(opt);
      });
    }
  } catch (err) {
    console.error('Failed to load templates for generator', err);
  }
}

async function handleDocxUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  uploadMultiFiles([file]);
}

async function uploadDocxFile(file) {
  uploadMultiFiles([file]);
}

// Multi-Source Input Handling
async function handleMultiSourceUpload(e) {
  const files = e.target.files;
  if (!files || files.length === 0) return;
  uploadMultiFiles(Array.from(files));
}

async function uploadMultiFiles(filesList) {
  if (!filesList || filesList.length === 0) return;

  const status = document.getElementById('upload-status');
  if (status) {
    status.classList.remove('hidden');
    status.innerText = `Uploading ${filesList.length} source file(s)...`;
    status.className = 'text-[11px] text-muted-foreground mt-1.5';
  }

  const formData = new FormData();
  filesList.forEach(f => formData.append('files', f));

  try {
    const res = await fetch('/api/generator/upload-multi', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      const newFiles = data.files || [];
      newFiles.forEach(nf => {
        if (!genSourceFiles.some(existing => existing.file_path === nf.file_path)) {
          genSourceFiles.push(nf);
        }
      });

      renderGenSourceFiles();

      if (genSourceFiles.length > 0) {
        const docxInput = document.getElementById('gen-docx-path');
        if (docxInput) docxInput.value = genSourceFiles[0].file_path;
        const outInput = document.getElementById('gen-output-path');
        if (outInput && !outInput.value) {
          outInput.value = data.suggested_output;
        }
      }

      if (status) {
        status.innerHTML = `<span class="text-emerald-400 font-medium">✓ Uploaded ${newFiles.length} file(s) successfully</span>`;
      }
      showToast(`Added ${newFiles.length} source file(s)`, 'success');
    } else {
      if (status) status.innerHTML = `<span class="text-destructive">Upload failed: ${data.error}</span>`;
      showToast(`Upload failed: ${data.error}`, 'error');
    }
  } catch (err) {
    if (status) status.innerHTML = `<span class="text-destructive">Upload error: ${err.message}</span>`;
    showToast(`Upload error: ${err.message}`, 'error');
  }
}

function renderGenSourceFiles() {
  const container = document.getElementById('gen-source-files-list');
  const countBadge = document.getElementById('gen-source-count-badge');
  if (!container) return;

  if (countBadge) countBadge.innerText = `${genSourceFiles.length} file(s)`;

  if (genSourceFiles.length === 0) {
    container.classList.add('hidden');
    container.innerHTML = '';
    return;
  }

  container.classList.remove('hidden');
  container.innerHTML = '';

  genSourceFiles.forEach((file, idx) => {
    const pill = document.createElement('div');
    pill.className = 'flex items-center justify-between p-2 rounded bg-secondary/70 border border-border text-xs gap-2';

    let iconName = 'file';
    let iconColor = 'text-primary';
    const cat = file.category || '';
    if (cat.includes('Word')) {
      iconName = 'file-text';
      iconColor = 'text-sky-400';
    } else if (cat.includes('PowerPoint')) {
      iconName = 'presentation';
      iconColor = 'text-amber-400';
    } else if (cat.includes('Image')) {
      iconName = 'image';
      iconColor = 'text-emerald-400';
    } else if (cat.includes('Audio')) {
      iconName = 'mic';
      iconColor = 'text-purple-400';
    } else if (cat.includes('Text')) {
      iconName = 'file-code';
      iconColor = 'text-cyan-400';
    }

    pill.innerHTML = `
      <div class="flex items-center gap-2 truncate flex-1 min-w-0">
        <i data-lucide="${iconName}" class="w-4 h-4 ${iconColor} flex-shrink-0"></i>
        <div class="truncate">
          <div class="font-medium text-foreground truncate text-[11px]">${file.filename}</div>
          <div class="text-[10px] text-muted-foreground font-mono">${file.category} • ${file.size_kb} KB</div>
        </div>
      </div>
      <button type="button" onclick="removeGenSourceFile(${idx})" class="text-muted-foreground hover:text-destructive p-1 rounded hover:bg-secondary flex-shrink-0" title="Remove file">
        <i data-lucide="x" class="w-3.5 h-3.5"></i>
      </button>
    `;
    container.appendChild(pill);
  });

  refreshIcons();
}

function removeGenSourceFile(index) {
  if (index >= 0 && index < genSourceFiles.length) {
    const removed = genSourceFiles.splice(index, 1)[0];
    renderGenSourceFiles();
    const docxInput = document.getElementById('gen-docx-path');
    if (docxInput) {
      docxInput.value = genSourceFiles.length > 0 ? genSourceFiles[0].file_path : '';
    }
    showToast(`Removed ${removed.filename}`, 'info', 1500);
  }
}

function toggleRawNotesInput() {
  const input = document.getElementById('gen-raw-text-input');
  const btnTxt = document.getElementById('toggle-notes-btn-text');
  if (!input) return;
  if (input.classList.contains('hidden')) {
    input.classList.remove('hidden');
    input.focus();
    if (btnTxt) btnTxt.innerText = '- Hide Direct Text Notes';
  } else {
    input.classList.add('hidden');
    if (btnTxt) btnTxt.innerText = '+ Paste Direct Text Notes / Speech Outline';
  }
}

// -------------------------------------------------------------
// STORYBOARD SPECIFICATIONS (data/structure/*.md)
// -------------------------------------------------------------
let currentStructureMode = 'unified';

function populateStructureSelect(selectEl, emptyLabel, selectedVal) {
  if (!selectEl) return;
  const prevVal = selectedVal !== undefined ? selectedVal : selectEl.value;
  selectEl.innerHTML = `<option value="">${emptyLabel}</option>`;
  structuresList.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s.filename;
    opt.innerText = `${s.filename} ${s.is_sample ? '★ (Sample Blueprint)' : ''} [${s.size}]`;
    selectEl.appendChild(opt);
  });
  if (prevVal && structuresList.some(s => s.filename === prevVal)) {
    selectEl.value = prevVal;
  }
}

async function loadGeneratorStructures() {
  try {
    const res = await fetch('/api/structure/list');
    const data = await res.json();
    structuresList = data.structures || [];

    populateStructureSelect(
      document.getElementById('gen-structure-select'),
      '✨ None (Standard Direct Generation - No Schema Restructuring)'
    );
    populateStructureSelect(
      document.getElementById('gen-detect-structure-select'),
      '✨ Auto / None (Standard Extraction)'
    );
    populateStructureSelect(
      document.getElementById('gen-restructure-structure-select'),
      '✨ None (Keep Original Document Sections)'
    );
    populateStructureSelect(
      document.getElementById('gen-blueprint-structure-select'),
      '✨ Inherit from Restructure / Detect Schema'
    );

    updateStructuresListUI();
    if (window.lucide) lucide.createIcons();
  } catch (err) {
    console.error('Failed to load structures list', err);
  }
}

function setStructureSelectionMode(mode) {
  currentStructureMode = mode;
  const unifiedBox = document.getElementById('gen-struct-unified-box');
  const dedicatedBox = document.getElementById('gen-struct-dedicated-box');
  const btnUnified = document.getElementById('mode-btn-unified');
  const btnDedicated = document.getElementById('mode-btn-dedicated');

  if (mode === 'dedicated') {
    if (unifiedBox) unifiedBox.classList.add('hidden');
    if (dedicatedBox) dedicatedBox.classList.remove('hidden');

    if (btnUnified) {
      btnUnified.className = 'flex-1 py-1 px-2 rounded font-medium text-center text-muted-foreground hover:text-foreground transition';
    }
    if (btnDedicated) {
      btnDedicated.className = 'flex-1 py-1 px-2 rounded font-medium text-center transition bg-primary text-primary-foreground shadow-sm';
    }

    const unifiedVal = document.getElementById('gen-structure-select')?.value;
    if (unifiedVal) {
      const dSel = document.getElementById('gen-detect-structure-select');
      const rSel = document.getElementById('gen-restructure-structure-select');
      if (dSel && !dSel.value) dSel.value = unifiedVal;
      if (rSel && !rSel.value) rSel.value = unifiedVal;
    }
  } else {
    if (unifiedBox) unifiedBox.classList.remove('hidden');
    if (dedicatedBox) dedicatedBox.classList.add('hidden');

    if (btnUnified) {
      btnUnified.className = 'flex-1 py-1 px-2 rounded font-medium text-center transition bg-primary text-primary-foreground shadow-sm';
    }
    if (btnDedicated) {
      btnDedicated.className = 'flex-1 py-1 px-2 rounded font-medium text-center text-muted-foreground hover:text-foreground transition';
    }
  }
  if (window.lucide) lucide.createIcons();
}

function onGeneratorStructureChange() {
  const sel = document.getElementById('gen-structure-select');
  const chkDetect = document.getElementById('gen-enable-detect');
  const chkRestruct = document.getElementById('gen-enable-restructure');
  const chkBlueprint = document.getElementById('gen-enable-blueprint');
  if (!sel) return;

  if (sel.value) {
    if (chkDetect) chkDetect.checked = true;
    if (chkRestruct) chkRestruct.checked = true;
    if (chkBlueprint) chkBlueprint.checked = true;
    showToast(`Selected structure blueprint: ${sel.value}. Restructure Agent enabled.`, 'info');
  } else {
    if (chkRestruct) chkRestruct.checked = false;
  }
}

function onDedicatedStructureChange(phase) {
  if (phase === 'detect') {
    const sel = document.getElementById('gen-detect-structure-select');
    const chk = document.getElementById('gen-detect-dedicated-enable');
    if (chk) chk.checked = Boolean(sel && sel.value);
    if (sel && sel.value) showToast(`Detection schema: ${sel.value}`, 'info');
  } else if (phase === 'restructure') {
    const sel = document.getElementById('gen-restructure-structure-select');
    const chk = document.getElementById('gen-restructure-dedicated-enable');
    if (chk) chk.checked = Boolean(sel && sel.value);
    if (sel && sel.value) showToast(`Restructure schema: ${sel.value}`, 'info');
  } else if (phase === 'blueprint') {
    const sel = document.getElementById('gen-blueprint-structure-select');
    const chk = document.getElementById('gen-blueprint-dedicated-enable');
    if (chk) chk.checked = Boolean(sel && sel.value);
    if (sel && sel.value) showToast(`Blueprint schema: ${sel.value}`, 'info');
  }
}

function triggerUploadStructure() {
  const inp = document.getElementById('gen-upload-structure-input');
  if (inp) inp.click();
}

async function handleUploadStructureFile(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const fd = new FormData();
  fd.append('file', file);

  try {
    showToast(`Importing structure schema: ${file.name}...`, 'info');
    const res = await fetch('/api/structure/upload', {
      method: 'POST',
      body: fd
    });
    const data = await res.json();
    if (data.success && data.uploaded) {
      const uploadedName = data.uploaded.filename;
      await loadGeneratorStructures();

      const uniSelect = document.getElementById('gen-structure-select');
      if (uniSelect) uniSelect.value = uploadedName;
      const detSelect = document.getElementById('gen-detect-structure-select');
      if (detSelect) detSelect.value = uploadedName;
      const restSelect = document.getElementById('gen-restructure-structure-select');
      if (restSelect) restSelect.value = uploadedName;

      onGeneratorStructureChange();
      showToast(`Successfully imported structure: ${uploadedName}`, 'success');
    } else {
      showToast(data.error || 'Failed to upload structure file', 'error');
    }
  } catch (err) {
    showToast(`Upload error: ${err.message}`, 'error');
  } finally {
    event.target.value = '';
  }
}

async function previewCurrentSelectedStructure(targetType = 'unified') {
  let selectId = 'gen-structure-select';
  if (targetType === 'detect') selectId = 'gen-detect-structure-select';
  else if (targetType === 'restructure') selectId = 'gen-restructure-structure-select';
  else if (targetType === 'blueprint') selectId = 'gen-blueprint-structure-select';

  const sel = document.getElementById(selectId);
  const name = sel ? sel.value : null;
  if (!name) {
    showToast(`Please select a ${targetType} schema file from the dropdown to view.`, 'warning');
    return;
  }
  openPreviewStructureModal(name);
}

async function openPreviewStructureModal(name) {
  const modal = document.getElementById('preview-structure-modal');
  const title = document.getElementById('preview-struct-modal-title');
  const pathEl = document.getElementById('preview-struct-modal-path');
  const contentEl = document.getElementById('preview-struct-content');
  if (!modal) return;

  if (contentEl) contentEl.innerText = 'Loading specification...';
  modal.classList.remove('hidden');

  try {
    const res = await fetch(`/api/structure/get?name=${encodeURIComponent(name)}`);
    const data = await res.json();
    if (data.success) {
      if (title) title.innerText = `Storyboard Schema: ${name}`;
      if (pathEl) pathEl.innerText = `data/structure/${name}`;
      if (contentEl) contentEl.innerText = data.content;
    } else {
      if (contentEl) contentEl.innerText = `Error loading structure: ${data.error}`;
    }
  } catch (err) {
    if (contentEl) contentEl.innerText = `Network error: ${err.message}`;
  }
  refreshIcons();
}

function closePreviewStructureModal() {
  const modal = document.getElementById('preview-structure-modal');
  if (modal) modal.classList.add('hidden');
}

function updateStructuresListUI() {
  const container = document.getElementById('structures-list-container');
  const badge = document.getElementById('struct-count-badge');
  if (!container) return;

  if (badge) badge.innerText = `${structuresList.length} files`;

  if (structuresList.length === 0) {
    container.innerHTML = '<div class="text-muted-foreground italic text-xs py-3">No structure specifications found in data/structure/.</div>';
    return;
  }

  container.innerHTML = '';
  structuresList.forEach(s => {
    const item = document.createElement('div');
    const isSelected = currentStructure && currentStructure.filename === s.filename;
    item.className = `p-2.5 rounded-md border cursor-pointer transition text-xs space-y-1 ${isSelected ? 'bg-secondary border-primary/70 font-medium' : 'bg-card/60 border-border hover:border-border/90 hover:bg-muted/40'}`;
    item.onclick = () => selectStructureItem(s.filename);

    item.innerHTML = `
      <div class="flex items-center justify-between">
        <span class="font-mono text-foreground font-semibold truncate max-w-[190px]">${s.filename}</span>
        ${s.is_sample ? '<span class="shadcn-badge shadcn-badge-outline text-amber-400 border-amber-800 text-[9px] py-0 px-1">Sample</span>' : ''}
      </div>
      <div class="text-[10px] text-muted-foreground flex items-center justify-between">
        <span>${s.size}</span>
        <span>${s.modified}</span>
      </div>
      ${s.excerpt ? `<div class="text-[11px] text-muted-foreground line-clamp-1 italic">${escapeHtml(s.excerpt)}</div>` : ''}
    `;
    container.appendChild(item);
  });

  if (!currentStructure && structuresList.length > 0) {
    selectStructureItem(structuresList[0].filename);
  }
}

async function selectStructureItem(filename) {
  try {
    const res = await fetch(`/api/structure/get?name=${encodeURIComponent(filename)}`);
    const data = await res.json();
    if (data.success) {
      currentStructure = { filename: filename, content: data.content };
      const nameInput = document.getElementById('struct-editor-filename');
      const contentInput = document.getElementById('struct-editor-content');
      if (nameInput) nameInput.value = filename;
      if (contentInput) contentInput.value = data.content;

      const delBtn = document.getElementById('btn-delete-current-structure');
      if (delBtn) {
        delBtn.disabled = (filename === 'pptx-structure-yosefzadeh.md');
      }

      updateStructuresListUI();
    }
  } catch (err) {
    console.error('Failed to get structure item', err);
  }
}

async function saveCurrentStructure() {
  let name = document.getElementById('struct-editor-filename')?.value.trim();
  const content = document.getElementById('struct-editor-content')?.value || '';

  if (!name) {
    showToast('Please provide a filename for the structure.', 'warning');
    return;
  }
  if (!name.endsWith('.md')) {
    name += '.md';
    document.getElementById('struct-editor-filename').value = name;
  }

  try {
    const res = await fetch('/api/structure/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, content })
    });
    const data = await res.json();
    if (data.success) {
      showToast(`Saved structure: ${name}`, 'success');
      await loadGeneratorStructures();
      selectStructureItem(name);
    } else {
      showToast(`Save failed: ${data.error}`, 'error');
    }
  } catch (err) {
    showToast(`Save error: ${err.message}`, 'error');
  }
}

async function deleteCurrentStructure() {
  const name = document.getElementById('struct-editor-filename')?.value.trim();
  if (!name) return;
  if (name === 'pptx-structure-yosefzadeh.md') {
    showToast('The sample reference structure cannot be deleted.', 'warning');
    return;
  }
  if (!confirm(`Are you sure you want to delete structure file: ${name}?`)) return;

  try {
    const res = await fetch('/api/structure/delete', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name })
    });
    const data = await res.json();
    if (data.success) {
      showToast(`Deleted ${name}`, 'info');
      currentStructure = null;
      await loadGeneratorStructures();
    } else {
      showToast(`Delete failed: ${data.error}`, 'error');
    }
  } catch (err) {
    showToast(`Delete error: ${err.message}`, 'error');
  }
}

function newStructureFile() {
  currentStructure = null;
  const defName = `structure-${Date.now().toString().slice(-4)}.md`;
  const nameInput = document.getElementById('struct-editor-filename');
  const contentInput = document.getElementById('struct-editor-content');
  if (nameInput) nameInput.value = defName;
  if (contentInput) {
    contentInput.value = `# Slide Storyboard Specification Schema (${defName})\n\n## 1. Document Metadata\n- **Subject / Topic**: [Topic]\n- **Slide_Range**: [Slides 1–4]\n\n---\n\n## 2. Quadrant Layout Mapping\n- Quadrant_TR: Slide 1 (Theory / Definition)\n- Quadrant_TL: Slide 2 (Classification)\n- Quadrant_BR: Slide 3 (Worked Example)\n- Quadrant_BL: Slide 4 (Exercise & Summary)\n\n---\n\n## 3. Slide Content Schema\n### Slide 1: [Title]\n- **Quadrant**: TR\n- **Slide_Type**: Theory / Definition\n- **Core_Concept**: [Brief 1-line concept]\n\n#### Animation Sequence (Numbered Steps)\n1. **Step 1 (①)**: [Text / Formula / Statement]\n2. **Step 2 (②)**: [Text / Formula / Statement]\n\n#### Mathematical Elements\n- **Formulas / Equations**: $E = mc^2$\n\n#### Visual Annotations\n- **Teacher Callouts**: اگه دقت کنی! ...\n`;
    contentInput.focus();
  }
  updateStructuresListUI();
}

function openBuildStructureModal() {
  const modal = document.getElementById('build-structure-modal');
  const tplSelect = document.getElementById('modal-struct-template-select');
  const nameInput = document.getElementById('modal-struct-name-input');
  const progBox = document.getElementById('modal-struct-progress');

  if (progBox) progBox.classList.add('hidden');

  if (tplSelect && templatesList.length > 0) {
    tplSelect.innerHTML = '<option value="">Select target template...</option>';
    templatesList.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t.filename;
      opt.innerText = `${t.filename} (${t.slide_count} slides) - ${t.style}`;
      tplSelect.appendChild(opt);
    });

    if (selectedTemplateName) {
      tplSelect.value = selectedTemplateName;
    } else {
      tplSelect.value = templatesList[0].filename;
    }
  }

  const chosenTpl = tplSelect?.value || 'template';
  const cleanStem = chosenTpl.replace(/\.pptx$/i, '');
  if (nameInput) {
    nameInput.value = `${cleanStem}-structure`;
  }

  modal.classList.remove('hidden');
  refreshIcons();
}

function closeBuildStructureModal() {
  const modal = document.getElementById('build-structure-modal');
  if (modal) modal.classList.add('hidden');
}

async function submitBuildStructure() {
  const tplName = document.getElementById('modal-struct-template-select')?.value;
  let structName = document.getElementById('modal-struct-name-input')?.value.trim();
  const instructions = document.getElementById('modal-struct-instructions-input')?.value.trim();

  if (!tplName) {
    showToast('Please select a reference template.', 'warning');
    return;
  }
  if (!structName) {
    structName = `${tplName.replace(/\.pptx$/i, '')}-structure`;
  }
  if (structName.endsWith('.md')) {
    structName = structName.replace(/\.md$/i, '');
  }

  const btnSubmit = document.getElementById('btn-submit-build-structure');
  const progBox = document.getElementById('modal-struct-progress');
  const logLine = document.getElementById('modal-struct-log-line');

  btnSubmit.disabled = true;
  progBox.classList.remove('hidden');
  logLine.innerText = `Connecting to AI model to inspect ${tplName}...`;

  try {
    const res = await fetch('/api/structure/build', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        template_name: tplName,
        structure_name: structName,
        custom_instructions: instructions
      })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const evtSource = new EventSource(`/api/generator/stream/${data.job_id}`);
    evtSource.addEventListener('log', (e) => {
      const d = JSON.parse(e.data);
      if (logLine) logLine.innerText = d.message;
    });

    evtSource.addEventListener('completed', (e) => {
      const d = JSON.parse(e.data);
      evtSource.close();
      btnSubmit.disabled = false;
      closeBuildStructureModal();
      showToast(`Successfully created structure: ${d.filename}`, 'success');
      loadGeneratorStructures().then(() => {
        selectStructureItem(d.filename);
      });
    });

    evtSource.addEventListener('error', (e) => {
      evtSource.close();
      btnSubmit.disabled = false;
      logLine.innerText = 'Failed generating structure.';
      showToast('Error synthesizing structure.md', 'error');
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    btnSubmit.disabled = false;
    logLine.innerText = `Error: ${err.message}`;
    showToast(`Build error: ${err.message}`, 'error');
  }
}

// -------------------------------------------------------------
// BUILD DETECTION STRUCTURE FROM SAMPLES
// -------------------------------------------------------------
function openBuildStructureFromSamplesModal() {
  const modal = document.getElementById('build-structure-from-samples-modal');
  const progBox = document.getElementById('modal-sample-struct-progress');
  const nameInput = document.getElementById('modal-sample-struct-name-input');

  if (progBox) progBox.classList.add('hidden');
  if (nameInput && !nameInput.value) {
    nameInput.value = `paper-detection-schema-${Date.now().toString().slice(-4)}`;
  }

  structSampleFiles = [];
  renderStructSamplePills();

  if (modal) modal.classList.remove('hidden');
  refreshIcons();
}

function closeBuildStructureFromSamplesModal() {
  const modal = document.getElementById('build-structure-from-samples-modal');
  if (modal) modal.classList.add('hidden');
}

async function handleStructSamplesUpload(e) {
  const files = e.target.files;
  if (!files || files.length === 0) return;
  uploadStructSampleFiles(Array.from(files));
}

async function uploadStructSampleFiles(filesList) {
  if (!filesList || filesList.length === 0) return;

  const formData = new FormData();
  filesList.forEach(f => formData.append('files', f));

  try {
    const res = await fetch('/api/structure/upload-samples', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (data.success && data.files) {
      data.files.forEach(nf => {
        if (!structSampleFiles.some(existing => existing.file_path === nf.file_path)) {
          structSampleFiles.push(nf);
        }
      });
      renderStructSamplePills();
      showToast(`Uploaded ${data.files.length} sample file(s)`, 'success', 2000);
    } else {
      showToast(`Upload failed: ${data.error}`, 'error');
    }
  } catch (err) {
    showToast(`Upload error: ${err.message}`, 'error');
  }
}

function renderStructSamplePills() {
  const container = document.getElementById('struct-samples-pills-list');
  if (!container) return;

  if (structSampleFiles.length === 0) {
    container.classList.add('hidden');
    container.innerHTML = '';
    return;
  }

  container.classList.remove('hidden');
  container.innerHTML = '';

  structSampleFiles.forEach((file, idx) => {
    const pill = document.createElement('div');
    pill.className = 'flex items-center justify-between p-2 rounded bg-secondary/70 border border-border text-xs gap-2';

    let icon = 'file-image';
    if (file.category.includes('Word')) icon = 'file-text';
    else if (file.category.includes('Image')) icon = 'image';

    pill.innerHTML = `
      <div class="flex items-center gap-2 truncate flex-1 min-w-0">
        <i data-lucide="${icon}" class="w-4 h-4 text-primary flex-shrink-0"></i>
        <div class="truncate">
          <div class="font-medium text-foreground truncate text-[11px]">${file.filename}</div>
          <div class="text-[10px] text-muted-foreground font-mono">${file.category} • ${file.size_kb} KB</div>
        </div>
      </div>
      <button type="button" onclick="removeStructSamplePill(${idx})" class="text-muted-foreground hover:text-destructive p-1 rounded hover:bg-secondary flex-shrink-0">
        <i data-lucide="x" class="w-3.5 h-3.5"></i>
      </button>
    `;
    container.appendChild(pill);
  });

  refreshIcons();
}

function removeStructSamplePill(idx) {
  if (idx >= 0 && idx < structSampleFiles.length) {
    structSampleFiles.splice(idx, 1);
    renderStructSamplePills();
  }
}

async function submitBuildStructureFromSamples() {
  if (structSampleFiles.length === 0) {
    showToast('Please upload at least one sample file or photo.', 'warning');
    return;
  }

  let structName = document.getElementById('modal-sample-struct-name-input')?.value.trim();
  const instructions = document.getElementById('modal-sample-struct-instructions-input')?.value.trim();

  if (!structName) {
    const firstStem = structSampleFiles[0].filename.split('.')[0];
    structName = `${firstStem}-detection-schema`;
  }
  if (structName.endsWith('.md')) {
    structName = structName.replace(/\.md$/i, '');
  }

  const btnSubmit = document.getElementById('btn-submit-build-sample-struct');
  const progBox = document.getElementById('modal-sample-struct-progress');
  const logLine = document.getElementById('modal-sample-struct-log-line');

  btnSubmit.disabled = true;
  progBox.classList.remove('hidden');
  logLine.innerText = `Connecting to Vision AI to analyze ${structSampleFiles.length} sample(s)...`;

  try {
    const res = await fetch('/api/structure/build-from-samples', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sample_files: structSampleFiles.map(f => f.file_path),
        structure_name: structName,
        custom_instructions: instructions
      })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const evtSource = new EventSource(`/api/generator/stream/${data.job_id}`);
    evtSource.addEventListener('log', (e) => {
      const d = JSON.parse(e.data);
      if (logLine) logLine.innerText = d.message;
    });

    evtSource.addEventListener('completed', (e) => {
      const d = JSON.parse(e.data);
      evtSource.close();
      btnSubmit.disabled = false;
      closeBuildStructureFromSamplesModal();
      showToast(`Successfully created detection structure: ${d.filename}`, 'success');
      loadGeneratorStructures().then(() => {
        selectStructureItem(d.filename);
      });
    });

    evtSource.addEventListener('error', (e) => {
      evtSource.close();
      btnSubmit.disabled = false;
      logLine.innerText = 'Failed generating detection structure.';
      showToast('Error synthesizing detection schema', 'error');
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    btnSubmit.disabled = false;
    logLine.innerText = `Error: ${err.message}`;
    showToast(`Build error: ${err.message}`, 'error');
  }
}

function appendGenLog(msg, type = 'info') {
  const logs = document.getElementById('gen-console-logs');
  const line = document.createElement('div');

  if (msg.startsWith('[*]')) {
    line.className = 'text-sky-400 font-medium';
  } else if (msg.startsWith('[✓]') || msg.includes('SUCCESS') || msg.includes('complete')) {
    line.className = 'text-emerald-400 font-medium';
  } else if (msg.startsWith('[!]') || msg.includes('ERROR') || msg.includes('failed')) {
    line.className = 'text-destructive font-semibold';
  } else {
    line.className = 'text-foreground/80';
  }

  line.innerText = msg;
  logs.appendChild(line);
  logs.scrollTop = logs.scrollHeight;
}

function clearGenLog() {
  document.getElementById('gen-console-logs').innerHTML = '';
}

function copyGenLogsToClipboard() {
  const logs = document.getElementById('gen-console-logs').innerText;
  if (!logs) return;
  navigator.clipboard.writeText(logs).then(() => {
    showToast('Execution logs copied to clipboard', 'success');
  }).catch(() => {
    showToast('Failed to copy logs', 'error');
  });
}

async function startPresentationGeneration() {
  const sourceFiles = genSourceFiles.map(f => f.file_path);
  const docxPath = document.getElementById('gen-docx-path')?.value.trim() || '';
  const rawText = document.getElementById('gen-raw-text-input')?.value.trim() || '';
  const templateName = document.getElementById('gen-template-select')?.value;
  let structureName = '';
  let detectionStructureName = '';
  let restructureStructureName = '';
  let blueprintStructureName = '';
  let enableDetect = true;
  let enableRestructure = false;
  let enableBlueprint = true;

  if (currentStructureMode === 'dedicated') {
    detectionStructureName = document.getElementById('gen-detect-structure-select')?.value || '';
    restructureStructureName = document.getElementById('gen-restructure-structure-select')?.value || '';
    blueprintStructureName = document.getElementById('gen-blueprint-structure-select')?.value || '';
    enableDetect = document.getElementById('gen-detect-dedicated-enable')?.checked ?? Boolean(detectionStructureName);
    enableRestructure = document.getElementById('gen-restructure-dedicated-enable')?.checked ?? Boolean(restructureStructureName);
    enableBlueprint = document.getElementById('gen-blueprint-dedicated-enable')?.checked ?? Boolean(blueprintStructureName);
    structureName = restructureStructureName || detectionStructureName || blueprintStructureName;
  } else {
    structureName = document.getElementById('gen-structure-select')?.value || '';
    enableDetect = document.getElementById('gen-enable-detect')?.checked ?? true;
    enableRestructure = document.getElementById('gen-enable-restructure')?.checked ?? false;
    enableBlueprint = document.getElementById('gen-enable-blueprint')?.checked ?? true;
  }

  const outputPath = document.getElementById('gen-output-path')?.value.trim();

  if (sourceFiles.length === 0 && !docxPath && !rawText) {
    showToast('Please upload source files (Word, PPTX, Text, Images, Audio) or enter text notes.', 'warning');
    return;
  }

  const btnGen = document.getElementById('btn-generate-pptx');
  const btnOpen = document.getElementById('btn-open-ppt');
  const btnDownload = document.getElementById('btn-download-pptx');

  btnGen.disabled = true;
  btnOpen.disabled = true;
  btnDownload.disabled = true;

  setSystemStatus('SYNTHESIZING...', true);
  const inputSummary = sourceFiles.length > 0 ? `${sourceFiles.length} source file(s)` : (docxPath ? docxPath.split(/[\\/]/).pop() : 'Direct notes');
  appendGenLog(`\n[*] Starting presentation synthesis for ${inputSummary}`);

  if (currentStructureMode === 'dedicated') {
    if (detectionStructureName && enableDetect) {
      appendGenLog(`[*] Detection Schema ACTIVE: '${detectionStructureName}' (Quadrant OCR & math extraction)`);
    }
    if (restructureStructureName && enableRestructure) {
      appendGenLog(`[*] AI Storyboard Restructure Agent ACTIVE: Conforming to '${restructureStructureName}'`);
    }
    if (blueprintStructureName && enableBlueprint) {
      appendGenLog(`[*] Slide Blueprint ACTIVE: Conforming to '${blueprintStructureName}'`);
    }
  } else if (structureName) {
    if (enableDetect) appendGenLog(`[*] Detection Schema ACTIVE: '${structureName}'`);
    if (enableRestructure) appendGenLog(`[*] AI Storyboard Restructure Agent ACTIVE: Conforming to '${structureName}'`);
    if (enableBlueprint) appendGenLog(`[*] Slide Blueprint ACTIVE: Conforming to '${structureName}'`);
  }

  const timeoutVal = parseInt(document.getElementById('gen-timeout-input')?.value || '300', 10);

  try {
    const res = await fetch('/api/generator/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        source_files: sourceFiles,
        docx_path: docxPath,
        raw_text: rawText,
        template_name: templateName,
        structure_name: structureName,
        detection_structure_name: detectionStructureName,
        restructure_structure_name: restructureStructureName,
        blueprint_structure_name: blueprintStructureName,
        enable_detection: enableDetect,
        enable_restructure: enableRestructure,
        enable_blueprint: enableBlueprint,
        output_path: outputPath,
        timeout: timeoutVal
      })
    });

    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || 'Failed to start generation job');
    }

    const jobId = data.job_id;
    listenToGenerationSSE(jobId);
  } catch (err) {
    appendGenLog(`[!] Generation failed: ${err.message}`);
    setSystemStatus('FAILED');
    btnGen.disabled = false;
    showToast(`Generation Error: ${err.message}`, 'error');
  }
}

function listenToGenerationSSE(jobId) {
  const evtSource = new EventSource(`/api/generator/stream/${jobId}`);
  const btnGen = document.getElementById('btn-generate-pptx');
  const btnOpen = document.getElementById('btn-open-ppt');
  const btnDownload = document.getElementById('btn-download-pptx');

  evtSource.addEventListener('log', (e) => {
    const d = JSON.parse(e.data);
    appendGenLog(d.message);
  });

  evtSource.addEventListener('ai_images', (e) => {
    const d = JSON.parse(e.data);
    aiTestImages = d.images || [];
    aiTestIdx = 0;
    updateAiTestDisplay();
  });

  evtSource.addEventListener('completed', (e) => {
    const d = JSON.parse(e.data);
    currentGeneratedPptx = d.pptx_path;
    genSlides = d.previews || [];
    genSlideIdx = 0;
    visualSlides = [...genSlides];
    visualSlideIdx = 0;

    updateGenSlideDisplay(d.engine_name);
    updateVisualSlideDisplay();
    renderSlideThumbnails();

    btnOpen.disabled = false;
    btnDownload.disabled = false;
    btnGen.disabled = false;

    setSystemStatus('READY');
    appendGenLog(`[✓] Completed: ${d.filename}`, 'success');
    showToast(`Generation complete: ${d.filename}`, 'success');
  });

  evtSource.addEventListener('error', (e) => {
    try {
      const d = JSON.parse(e.data);
      appendGenLog(`[!] Error: ${d.error}`);
    } catch (_) {}
    setSystemStatus('ERROR');
    btnGen.disabled = false;
    showToast('Slide generation failed.', 'error');
  });

  evtSource.addEventListener('close', () => {
    evtSource.close();
  });
}

function updateGenSlideDisplay(engineName = '') {
  const img = document.getElementById('gen-slide-img');
  const ph = document.getElementById('gen-slide-placeholder');
  const counter = document.getElementById('gen-slide-counter');
  const badge = document.getElementById('gen-engine-badge');
  const prevBtn = document.getElementById('gen-prev-btn');
  const nextBtn = document.getElementById('gen-next-btn');

  if (!genSlides || genSlides.length === 0) {
    img.classList.add('hidden');
    ph.classList.remove('hidden');
    counter.innerText = 'No slides loaded';
    badge.classList.add('hidden');
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = genSlides[genSlideIdx].data_url;

  counter.innerText = `Slide ${genSlideIdx + 1} of ${genSlides.length}`;
  prevBtn.disabled = genSlideIdx === 0;
  nextBtn.disabled = genSlideIdx === genSlides.length - 1;

  if (engineName) {
    badge.classList.remove('hidden');
    if (engineName.includes('PowerPoint')) {
      badge.innerText = 'Native PowerPoint';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 border border-emerald-800 text-emerald-300';
    } else if (engineName.includes('Web')) {
      badge.innerText = 'Web Render Engine';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-sky-950 border border-sky-800 text-sky-300';
    } else {
      badge.innerText = 'Pure PIL';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-secondary border border-border text-foreground';
    }
  }

  updateActiveThumbnail();
}

function renderSlideThumbnails() {
  const container = document.getElementById('gen-thumbnails-strip');
  const countBadge = document.getElementById('gen-thumb-count');
  if (!container) return;

  if (!genSlides || genSlides.length === 0) {
    container.innerHTML = '<div class="text-[11px] text-muted-foreground italic px-1">Thumbnails will appear upon generation.</div>';
    if (countBadge) countBadge.innerText = '0 slides';
    return;
  }

  if (countBadge) countBadge.innerText = `${genSlides.length} slides`;
  container.innerHTML = '';

  genSlides.forEach((slide, idx) => {
    const thumb = document.createElement('div');
    thumb.className = `flex-shrink-0 w-20 h-12 rounded border border-border cursor-pointer bg-white transition relative ${idx === genSlideIdx ? 'thumb-active' : 'opacity-70 hover:opacity-100'}`;
    thumb.onclick = () => {
      genSlideIdx = idx;
      updateGenSlideDisplay();
    };

    thumb.innerHTML = `
      <img src="${slide.data_url}" alt="Slide ${idx + 1}" class="w-full h-full object-contain rounded">
      <div class="absolute bottom-0.5 right-0.5 bg-black/80 text-[8px] font-mono text-white px-1 rounded">${idx + 1}</div>
    `;
    container.appendChild(thumb);
  });
}

function updateActiveThumbnail() {
  const container = document.getElementById('gen-thumbnails-strip');
  if (!container) return;
  const thumbs = container.children;
  for (let i = 0; i < thumbs.length; i++) {
    if (i === genSlideIdx) {
      thumbs[i].className = 'flex-shrink-0 w-20 h-12 rounded border cursor-pointer bg-white transition relative thumb-active';
    } else {
      thumbs[i].className = 'flex-shrink-0 w-20 h-12 rounded border border-border cursor-pointer bg-white transition relative opacity-70 hover:opacity-100';
    }
  }
}

function navGenSlide(dir) {
  if (genSlideIdx + dir >= 0 && genSlideIdx + dir < genSlides.length) {
    genSlideIdx += dir;
    updateGenSlideDisplay();
  }
}

// Lightbox
function openCurrentSlideInLightbox() {
  if (!genSlides || genSlides.length === 0) return;
  lightboxSlides = genSlides;
  lightboxIdx = genSlideIdx;

  const modal = document.getElementById('slide-lightbox');
  const img = document.getElementById('lightbox-img');
  const counter = document.getElementById('lightbox-counter');

  img.src = lightboxSlides[lightboxIdx].data_url;
  counter.innerText = `${lightboxIdx + 1} / ${lightboxSlides.length}`;
  modal.classList.remove('hidden');
  refreshIcons();
}

function closeSlideLightbox() {
  const modal = document.getElementById('slide-lightbox');
  if (modal) modal.classList.add('hidden');
}

function navLightboxSlide(dir) {
  if (!lightboxSlides || lightboxSlides.length === 0) return;
  if (lightboxIdx + dir >= 0 && lightboxIdx + dir < lightboxSlides.length) {
    lightboxIdx += dir;
    const img = document.getElementById('lightbox-img');
    const counter = document.getElementById('lightbox-counter');
    img.src = lightboxSlides[lightboxIdx].data_url;
    counter.innerText = `${lightboxIdx + 1} / ${lightboxSlides.length}`;
  }
}

function updateVisualSlideDisplay() {
  const img = document.getElementById('visual-slide-img');
  const ph = document.getElementById('visual-slide-placeholder');
  const counter = document.getElementById('visual-slide-counter');
  const prevBtn = document.getElementById('visual-prev-btn');
  const nextBtn = document.getElementById('visual-next-btn');

  if (!visualSlides || visualSlides.length === 0) {
    img.classList.add('hidden');
    ph.classList.remove('hidden');
    counter.innerText = 'No screenshots loaded';
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = visualSlides[visualSlideIdx].data_url;
  counter.innerText = `Screenshot ${visualSlideIdx + 1} of ${visualSlides.length}`;
  prevBtn.disabled = visualSlideIdx === 0;
  nextBtn.disabled = visualSlideIdx === visualSlides.length - 1;
}

function navVisualSlide(dir) {
  if (visualSlideIdx + dir >= 0 && visualSlideIdx + dir < visualSlides.length) {
    visualSlideIdx += dir;
    updateVisualSlideDisplay();
  }
}

function updateAiTestDisplay() {
  const img = document.getElementById('ai-test-slide-img');
  const ph = document.getElementById('ai-test-slide-placeholder');
  const counter = document.getElementById('ai-test-slide-counter');
  const prevBtn = document.getElementById('ai-test-prev-btn');
  const nextBtn = document.getElementById('ai-test-next-btn');

  if (!aiTestImages || aiTestImages.length === 0) {
    img.classList.add('hidden');
    ph.classList.remove('hidden');
    counter.innerText = 'No AI payload images';
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  const cur = aiTestImages[aiTestIdx];
  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = cur.base64;
  counter.innerText = `Payload ${aiTestIdx + 1}/${aiTestImages.length} • ${cur.template_file} [Slide ${cur.slide_index + 1}]`;
  prevBtn.disabled = aiTestIdx === 0;
  nextBtn.disabled = aiTestIdx === aiTestImages.length - 1;
}

function navAiTestSlide(dir) {
  if (aiTestIdx + dir >= 0 && aiTestIdx + dir < aiTestImages.length) {
    aiTestIdx += dir;
    updateAiTestDisplay();
  }
}

async function openCurrentInPowerpoint() {
  if (!currentGeneratedPptx) return;
  try {
    await fetch('/api/manager/open', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: currentGeneratedPptx })
    });
    showToast('Launched presentation in PowerPoint', 'info');
  } catch (err) {
    showToast(`Could not launch presentation: ${err.message}`, 'error');
  }
}

function downloadCurrentPptx() {
  if (!currentGeneratedPptx) return;
  window.open(`/api/manager/download?file=${encodeURIComponent(currentGeneratedPptx)}`, '_blank');
}

// -------------------------------------------------------------
// 2. TEMPLATE INTELLIGENCE & ANALYZER
// -------------------------------------------------------------
async function loadTemplatesList() {
  try {
    const res = await fetch('/api/templates/list');
    const data = await res.json();
    templatesList = data.templates || [];

    document.getElementById('tpl-count-badge').innerText = `${data.analyzed_count}/${data.total_count} Analyzed`;

    const tbody = document.getElementById('templates-table-body');
    tbody.innerHTML = '';

    if (templatesList.length === 0) {
      tbody.innerHTML = '<tr><td colspan="4" class="py-4 text-center text-muted-foreground italic">No templates found in data/.</td></tr>';
      return;
    }

    templatesList.forEach((tpl) => {
      const tr = document.createElement('tr');
      tr.className = `cursor-pointer hover:bg-muted/50 transition ${selectedTemplateName === tpl.filename ? 'bg-muted font-medium' : ''}`;
      tr.onclick = () => selectTemplateItem(tpl.filename);

      const statusTag = tpl.is_analyzed
        ? '<span class="shadcn-badge shadcn-badge-outline text-emerald-400 border-emerald-800">✓ Analyzed</span>'
        : '<span class="shadcn-badge shadcn-badge-outline text-muted-foreground">○ Pending</span>';

      tr.innerHTML = `
        <td class="py-2 px-3 font-mono text-foreground">${tpl.filename}</td>
        <td class="py-2 px-3 text-center text-muted-foreground font-mono">${tpl.slide_count}</td>
        <td class="py-2 px-3 text-center">${statusTag}</td>
        <td class="py-2 px-3 text-muted-foreground truncate max-w-[140px]">${tpl.style} • ${tpl.purpose}</td>
      `;
      tbody.appendChild(tr);
    });

    if (!selectedTemplateName && templatesList.length > 0) {
      selectTemplateItem(templatesList[0].filename);
    }
  } catch (err) {
    console.error('Error loading template list', err);
  }
}

async function selectTemplateItem(filename) {
  selectedTemplateName = filename;
  loadTemplatesListVisuals();

  const tpl = templatesList.find(t => t.filename === filename);
  const metaBox = document.getElementById('tpl-detail-meta');

  if (tpl) {
    metaBox.innerHTML = `
      <div class="font-semibold text-foreground text-xs">${tpl.filename}</div>
      <div class="text-muted-foreground text-[11px] font-mono">Slides: ${tpl.slide_count} | Dim: ${tpl.dimensions}</div>
      <div class="pt-2 border-t border-border space-y-1">
        <div><span class="text-primary font-medium">🎯 Purpose:</span> ${tpl.purpose || 'Not analyzed'}</div>
        <div><span class="text-sky-400 font-medium">🎨 Style:</span> ${tpl.style || 'Not analyzed'}</div>
        <div><span class="text-amber-400 font-medium">📝 Brief:</span> ${tpl.brief || 'Click Analyze to generate'}</div>
      </div>
    `;

    loadTemplateSlidePreviews(tpl.file_path);
  }
}

function loadTemplatesListVisuals() {
  document.querySelectorAll('#templates-table-body tr').forEach(tr => {
    if (tr.innerText.includes(selectedTemplateName)) {
      tr.classList.add('bg-muted');
    } else {
      tr.classList.remove('bg-muted');
    }
  });
}

async function loadTemplateSlidePreviews(filePath) {
  const ph = document.getElementById('tpl-slide-placeholder');
  const img = document.getElementById('tpl-slide-img');

  ph.innerText = 'Rendering slides...';
  ph.classList.remove('hidden');
  img.classList.add('hidden');

  try {
    const res = await fetch('/api/preview/render', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: filePath, width: 450 })
    });
    const data = await res.json();
    if (data.success && data.slides && data.slides.length > 0) {
      tplSlides = data.slides;
      tplSlideIdx = 0;
      updateTplSlideDisplay();
    } else {
      ph.innerText = 'No preview available';
    }
  } catch (err) {
    ph.innerText = 'Render failed';
  }
}

function updateTplSlideDisplay() {
  const ph = document.getElementById('tpl-slide-placeholder');
  const img = document.getElementById('tpl-slide-img');
  const counter = document.getElementById('tpl-slide-counter');

  if (!tplSlides || tplSlides.length === 0) {
    img.classList.add('hidden');
    ph.classList.remove('hidden');
    counter.innerText = '0 / 0';
    return;
  }

  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = tplSlides[tplSlideIdx].data_url;
  counter.innerText = `${tplSlideIdx + 1} / ${tplSlides.length}`;
}

function navTplSlide(dir) {
  if (tplSlideIdx + dir >= 0 && tplSlideIdx + dir < tplSlides.length) {
    tplSlideIdx += dir;
    updateTplSlideDisplay();
  }
}

async function loadNoteMd() {
  try {
    const res = await fetch('/api/templates/notes');
    const data = await res.json();
    document.getElementById('note-md-editor').value = data.content || '';
  } catch (err) {
    console.error('Failed to load NOTE.md', err);
  }
}

async function saveNoteMd() {
  const content = document.getElementById('note-md-editor').value;
  try {
    const res = await fetch('/api/templates/notes', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ content })
    });
    const data = await res.json();
    if (data.success) {
      showToast('Knowledge Base (data/NOTE.md) saved successfully!', 'success');
      loadTemplatesList();
    }
  } catch (err) {
    showToast(`Save error: ${err.message}`, 'error');
  }
}

function appendAnalyzeLog(msg) {
  const logs = document.getElementById('analyze-console-logs');
  const line = document.createElement('div');
  line.innerText = msg;
  logs.appendChild(line);
  logs.scrollTop = logs.scrollHeight;
}

function clearAnalyzeLog() {
  document.getElementById('analyze-console-logs').innerHTML = '';
}

async function analyzeSelectedTemplate() {
  if (!selectedTemplateName) {
    showToast('Please select a template from the table first.', 'warning');
    return;
  }

  const btnSel = document.getElementById('btn-analyze-sel');
  const btnAll = document.getElementById('btn-analyze-all');
  btnSel.disabled = true;
  btnAll.disabled = true;

  appendAnalyzeLog(`\n[*] Starting AI analysis for template: ${selectedTemplateName}`);
  showToast(`Analyzing template: ${selectedTemplateName}`, 'info');

  try {
    const res = await fetch('/api/templates/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: selectedTemplateName })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const evtSource = new EventSource(`/api/generator/stream/${data.job_id}`);
    evtSource.addEventListener('log', (e) => {
      const d = JSON.parse(e.data);
      appendAnalyzeLog(d.message);
    });
    evtSource.addEventListener('completed', () => {
      appendAnalyzeLog(`[✓] Finished analysis for ${selectedTemplateName}`);
      loadTemplatesList();
      loadNoteMd();
      btnSel.disabled = false;
      btnAll.disabled = false;
      showToast(`Analysis completed for ${selectedTemplateName}`, 'success');
    });
    evtSource.addEventListener('error', () => {
      appendAnalyzeLog(`[!] Error analyzing template`);
      btnSel.disabled = false;
      btnAll.disabled = false;
      showToast('Template analysis error', 'error');
    });
    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAnalyzeLog(`[!] Analysis failed: ${err.message}`);
    btnSel.disabled = false;
    btnAll.disabled = false;
    showToast(`Analysis failed: ${err.message}`, 'error');
  }
}

async function analyzeAllTemplatesBatch() {
  if (!confirm('Analyze all templates with AI Agent? This will update data/NOTE.md with structured archetype knowledge.')) {
    return;
  }

  const btnSel = document.getElementById('btn-analyze-sel');
  const btnAll = document.getElementById('btn-analyze-all');
  const progContainer = document.getElementById('analyze-prog-container');
  const progBar = document.getElementById('analyze-prog-bar');
  const progStatus = document.getElementById('analyze-prog-status');
  const progPct = document.getElementById('analyze-prog-pct');

  btnSel.disabled = true;
  btnAll.disabled = true;
  progContainer.classList.remove('hidden');
  progBar.style.width = '0%';

  appendAnalyzeLog(`\n[*] Starting batch template analysis pipeline...`);
  showToast('Starting batch template analysis...', 'info');

  try {
    const res = await fetch('/api/templates/analyze-all', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const evtSource = new EventSource(`/api/generator/stream/${data.job_id}`);
    evtSource.addEventListener('log', (e) => {
      const d = JSON.parse(e.data);
      appendAnalyzeLog(d.message);
    });

    evtSource.addEventListener('progress', (e) => {
      const d = JSON.parse(e.data);
      progBar.style.width = `${d.percentage}%`;
      progPct.innerText = `${d.percentage}%`;
      progStatus.innerText = `Analyzing [${d.current}/${d.total}]: ${d.current_name}`;
    });

    evtSource.addEventListener('completed', () => {
      appendAnalyzeLog(`[✓] Completed batch analysis of all templates.`);
      progBar.style.width = '100%';
      progPct.innerText = '100%';
      loadTemplatesList();
      loadNoteMd();
      btnSel.disabled = false;
      btnAll.disabled = false;
      showToast('Batch template analysis complete!', 'success');
    });

    evtSource.addEventListener('error', () => {
      appendAnalyzeLog(`[!] Error in batch template analysis.`);
      btnSel.disabled = false;
      btnAll.disabled = false;
      showToast('Batch analysis failed', 'error');
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAnalyzeLog(`[!] Batch analysis failed: ${err.message}`);
    btnSel.disabled = false;
    btnAll.disabled = false;
    showToast(`Batch error: ${err.message}`, 'error');
  }
}

// -------------------------------------------------------------
// 3. DECK & TEMPLATE MANAGER
// -------------------------------------------------------------
async function loadManagerDecks() {
  try {
    const res = await fetch('/api/manager/decks');
    const data = await res.json();

    mgrGeneratedList = data.generated || [];
    mgrReferenceList = data.reference || [];

    document.getElementById('mgr-gen-count').innerText = `${mgrGeneratedList.length} decks`;
    document.getElementById('mgr-ref-count').innerText = `${mgrReferenceList.length} templates`;

    // Render Generated table
    const genBody = document.getElementById('mgr-gen-table-body');
    genBody.innerHTML = '';
    if (mgrGeneratedList.length === 0) {
      genBody.innerHTML = '<tr><td colspan="3" class="py-3 text-center text-muted-foreground italic">No generated decks found.</td></tr>';
    } else {
      mgrGeneratedList.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = `cursor-pointer hover:bg-muted/50 transition ${mgrSelectedFile?.file_path === item.file_path ? 'bg-muted font-medium' : ''}`;
        tr.onclick = () => selectManagerFile(item);
        tr.innerHTML = `
          <td class="py-2 px-3 font-mono text-foreground truncate max-w-[200px]">${item.filename}</td>
          <td class="py-2 px-3 text-center text-muted-foreground font-mono text-[11px]">${item.size}</td>
          <td class="py-2 px-3 text-center text-muted-foreground font-mono text-[11px]">${item.modified}</td>
        `;
        genBody.appendChild(tr);
      });
    }

    // Render Reference table
    const refBody = document.getElementById('mgr-ref-table-body');
    refBody.innerHTML = '';
    if (mgrReferenceList.length === 0) {
      refBody.innerHTML = '<tr><td colspan="3" class="py-3 text-center text-muted-foreground italic">No reference templates found.</td></tr>';
    } else {
      mgrReferenceList.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = `cursor-pointer hover:bg-muted/50 transition ${mgrSelectedFile?.file_path === item.file_path ? 'bg-muted font-medium' : ''}`;
        tr.onclick = () => selectManagerFile(item);
        tr.innerHTML = `
          <td class="py-2 px-3 font-mono text-foreground truncate max-w-[200px]">${item.filename}</td>
          <td class="py-2 px-3 text-center text-muted-foreground font-mono text-[11px]">${item.size}</td>
          <td class="py-2 px-3 text-center text-muted-foreground font-mono text-[11px]">${item.modified}</td>
        `;
        refBody.appendChild(tr);
      });
    }

    if (!mgrSelectedFile && mgrGeneratedList.length > 0) {
      selectManagerFile(mgrGeneratedList[0]);
    }
  } catch (err) {
    console.error('Error loading manager decks', err);
  }
}

async function handleTemplateImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/manager/upload-template', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      showToast(`Imported template: ${data.filename}`, 'success');
      loadManagerDecks();
    } else {
      showToast(`Import failed: ${data.error}`, 'error');
    }
  } catch (err) {
    showToast(`Import error: ${err.message}`, 'error');
  }
}

function selectManagerFile(item) {
  mgrSelectedFile = item;
  document.getElementById('mgr-sel-name').innerText = item.filename;
  document.getElementById('mgr-sel-details').innerText = `Type: ${item.type === 'generated' ? 'Generated Deck' : 'Reference Template'} | Size: ${item.size} | Modified: ${item.modified}\nPath: ${item.file_path}`;

  ['mgr-btn-open', 'mgr-btn-download', 'mgr-btn-verify', 'mgr-btn-duplicate', 'mgr-btn-rename', 'mgr-btn-delete'].forEach(id => {
    document.getElementById(id).disabled = false;
  });

  loadManagerSlidePreviews(item.file_path);
}

async function loadManagerSlidePreviews(filePath) {
  const ph = document.getElementById('mgr-slide-placeholder');
  const img = document.getElementById('mgr-slide-img');
  const counter = document.getElementById('mgr-slide-counter');
  const badge = document.getElementById('mgr-engine-badge');

  ph.innerText = 'Rendering slides...';
  ph.classList.remove('hidden');
  img.classList.add('hidden');
  badge.classList.add('hidden');

  try {
    const res = await fetch('/api/preview/render', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: filePath, width: 750 })
    });
    const data = await res.json();
    if (data.success && data.slides && data.slides.length > 0) {
      mgrSlides = data.slides;
      mgrSlideIdx = 0;
      updateMgrSlideDisplay(data.engine_name);
    } else {
      ph.innerText = 'No preview available';
    }
  } catch (err) {
    ph.innerText = 'Render failed';
  }
}

function updateMgrSlideDisplay(engineName = '') {
  const ph = document.getElementById('mgr-slide-placeholder');
  const img = document.getElementById('mgr-slide-img');
  const counter = document.getElementById('mgr-slide-counter');
  const prevBtn = document.getElementById('mgr-prev-btn');
  const nextBtn = document.getElementById('mgr-next-btn');
  const badge = document.getElementById('mgr-engine-badge');

  if (!mgrSlides || mgrSlides.length === 0) {
    img.classList.add('hidden');
    ph.classList.remove('hidden');
    counter.innerText = 'No slides loaded';
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = mgrSlides[mgrSlideIdx].data_url;
  counter.innerText = `Slide ${mgrSlideIdx + 1} of ${mgrSlides.length}`;
  prevBtn.disabled = mgrSlideIdx === 0;
  nextBtn.disabled = mgrSlideIdx === mgrSlides.length - 1;

  if (engineName) {
    badge.classList.remove('hidden');
    if (engineName.includes('PowerPoint')) {
      badge.innerText = 'Native PowerPoint';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 border border-emerald-800 text-emerald-300';
    } else if (engineName.includes('Web')) {
      badge.innerText = 'Web Render Engine';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-sky-950 border border-sky-800 text-sky-300';
    } else {
      badge.innerText = 'Pure PIL';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-secondary border border-border text-foreground';
    }
  }
}

function navMgrSlide(dir) {
  if (mgrSlideIdx + dir >= 0 && mgrSlideIdx + dir < mgrSlides.length) {
    mgrSlideIdx += dir;
    updateMgrSlideDisplay();
  }
}

async function mgrOpenPpt() {
  if (!mgrSelectedFile) return;
  try {
    await fetch('/api/manager/open', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: mgrSelectedFile.file_path })
    });
    showToast(`Opened: ${mgrSelectedFile.filename}`, 'info');
  } catch (err) {
    showToast(`Could not open file: ${err.message}`, 'error');
  }
}

function mgrDownload() {
  if (!mgrSelectedFile) return;
  window.open(`/api/manager/download?file=${encodeURIComponent(mgrSelectedFile.file_path)}`, '_blank');
}

async function mgrVerifyAndFix() {
  if (!mgrSelectedFile) return;
  try {
    const res = await fetch('/api/manager/verify', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: mgrSelectedFile.file_path })
    });
    const data = await res.json();
    if (data.success) {
      showToast(`Integrity check passed! Auto-healed: ${data.filename}`, 'success');
      loadManagerDecks();
    }
  } catch (err) {
    showToast(`Verify error: ${err.message}`, 'error');
  }
}

async function mgrDuplicate() {
  if (!mgrSelectedFile) return;
  try {
    const res = await fetch('/api/manager/duplicate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: mgrSelectedFile.file_path })
    });
    const data = await res.json();
    if (data.success) {
      showToast(`Created duplicate: ${data.filename}`, 'success');
      loadManagerDecks();
    }
  } catch (err) {
    showToast(`Duplicate error: ${err.message}`, 'error');
  }
}

async function mgrPromptRename() {
  if (!mgrSelectedFile) return;
  const newName = prompt(`Enter new presentation name:`, mgrSelectedFile.filename);
  if (!newName || newName === mgrSelectedFile.filename) return;

  try {
    const res = await fetch('/api/manager/rename', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_path: mgrSelectedFile.file_path,
        new_name: newName
      })
    });
    const data = await res.json();
    if (data.success) {
      showToast(`Renamed to: ${data.filename}`, 'success');
      loadManagerDecks();
    } else {
      showToast(`Rename failed: ${data.error}`, 'error');
    }
  } catch (err) {
    showToast(`Rename error: ${err.message}`, 'error');
  }
}

async function mgrDelete() {
  if (!mgrSelectedFile) return;
  if (!confirm(`Permanently delete presentation: ${mgrSelectedFile.filename}?`)) return;

  try {
    const res = await fetch('/api/manager/delete', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: mgrSelectedFile.file_path })
    });
    const data = await res.json();
    if (data.success) {
      showToast('Presentation deleted.', 'info');
      mgrSelectedFile = null;
      loadManagerDecks();
    }
  } catch (err) {
    showToast(`Delete error: ${err.message}`, 'error');
  }
}

// -------------------------------------------------------------
// 4. COMPONENT CATALOG (Visual Grid & Archetype Filter System)
// -------------------------------------------------------------
let componentsCatalogData = null;
let allComponentList = [];
let activeComponentFilter = 'all';
let currentModalComponent = null;

async function loadComponentsCatalog() {
  try {
    const res = await fetch('/api/components/catalog');
    const data = await res.json();
    componentsCatalogData = data.catalog || {};
    allComponentList = componentsCatalogData.all_components || [];

    document.getElementById('comp-count-badge').innerText = `Components: ${data.count || allComponentList.length}`;
    document.getElementById('components-json-viewer').value = JSON.stringify(componentsCatalogData, null, 2);

    updateFilterCounts();
    renderComponentCards();
  } catch (err) {
    console.error('Error loading component catalog', err);
  }
}

function switchComponentView(viewType) {
  const visualView = document.getElementById('comp-visual-view');
  const jsonView = document.getElementById('comp-json-view');
  const btnVisual = document.getElementById('btn-comp-view-visual');
  const btnJson = document.getElementById('btn-comp-view-json');

  if (viewType === 'visual') {
    visualView.classList.remove('hidden');
    jsonView.classList.add('hidden');
    btnVisual.classList.add('active', 'text-primary');
    btnJson.classList.remove('active', 'text-primary');
  } else {
    visualView.classList.add('hidden');
    jsonView.classList.remove('hidden');
    btnJson.classList.add('active', 'text-primary');
    btnVisual.classList.remove('active', 'text-primary');
  }
}

function updateFilterCounts() {
  const countAll = allComponentList.length;
  const countTitle = allComponentList.filter(c => matchCategory(c, 'title')).length;
  const countMetric = allComponentList.filter(c => matchCategory(c, 'metric')).length;
  const countCard = allComponentList.filter(c => matchCategory(c, 'card')).length;
  const countTable = allComponentList.filter(c => matchCategory(c, 'table')).length;
  const countImage = allComponentList.filter(c => matchCategory(c, 'image')).length;

  const elAll = document.getElementById('filter-count-all');
  const elTitle = document.getElementById('filter-count-title');
  const elMetric = document.getElementById('filter-count-metric');
  const elCard = document.getElementById('filter-count-card');
  const elTable = document.getElementById('filter-count-table');
  const elImage = document.getElementById('filter-count-image');

  if (elAll) elAll.innerText = countAll;
  if (elTitle) elTitle.innerText = countTitle;
  if (elMetric) elMetric.innerText = countMetric;
  if (elCard) elCard.innerText = countCard;
  if (elTable) elTable.innerText = countTable;
  if (elImage) elImage.innerText = countImage;
}

function matchCategory(comp, cat) {
  const label = (comp.label || '').toLowerCase();
  const desc = (comp.description || '').toLowerCase();
  const type = (comp.type || '').toLowerCase();

  if (cat === 'title') {
    return label.includes('title') || label.includes('header') || desc.includes('header') || desc.includes('title');
  } else if (cat === 'metric') {
    return label.includes('metric') || label.includes('stat') || desc.includes('numeric');
  } else if (cat === 'card') {
    return type.includes('shape') || label.includes('card') || label.includes('container') || label.includes('box');
  } else if (cat === 'table') {
    return type.includes('table') || label.includes('table');
  } else if (cat === 'image') {
    return type.includes('picture') || label.includes('image') || label.includes('graphic');
  }
  return true;
}

function setComponentFilter(filterName, btn) {
  activeComponentFilter = filterName;
  document.querySelectorAll('.component-filter-chip').forEach(c => c.classList.remove('active'));
  if (btn) btn.classList.add('active');
  renderComponentCards();
}

function filterComponentCards() {
  renderComponentCards();
}

function renderComponentCards() {
  const grid = document.getElementById('components-cards-grid');
  if (!grid) return;

  const query = (document.getElementById('comp-search-input')?.value || '').toLowerCase().trim();

  let filtered = allComponentList.filter(comp => {
    if (activeComponentFilter !== 'all' && !matchCategory(comp, activeComponentFilter)) {
      return false;
    }
    if (query) {
      const matchText = `${comp.label || ''} ${comp.description || ''} ${comp.sample_text || ''} ${comp.type || ''} ${comp.source_file || ''}`.toLowerCase();
      if (!matchText.includes(query)) {
        return false;
      }
    }
    return true;
  });

  if (filtered.length === 0) {
    grid.innerHTML = `
      <div class="col-span-full py-12 text-center text-muted-foreground text-xs">
        <i data-lucide="search-x" class="w-8 h-8 mx-auto mb-2 text-muted-foreground/60"></i>
        No matching components found for the selected filter or query.
      </div>
    `;
    refreshIcons();
    return;
  }

  grid.innerHTML = '';

  filtered.forEach(comp => {
    const card = document.createElement('div');
    card.className = 'component-card';
    card.onclick = () => openComponentModal(comp);

    // Render Preview Mini Block
    let previewMarkup = '';
    const sample = comp.sample_text || comp.label;
    const font = comp.font || {};
    const fill = comp.fill || {};

    let inlineStyle = '';
    if (font.name) inlineStyle += `font-family: ${font.name}, sans-serif; `;
    if (font.color) inlineStyle += `color: ${font.color}; `;
    if (font.bold) inlineStyle += `font-weight: 700; `;

    if (comp.type === 'PICTURE') {
      if (comp.image_path) {
        const imgName = comp.image_path.split(/[\\/]/).pop();
        previewMarkup = `<img src="/api/components/image/${encodeURIComponent(imgName)}" alt="${escapeHtml(comp.label)}" class="max-h-full max-w-full object-contain rounded">`;
      } else {
        previewMarkup = `<div class="flex items-center gap-1.5 text-muted-foreground"><i data-lucide="image" class="w-6 h-6 text-primary"></i> <span class="text-[11px] font-mono">Image Primitive</span></div>`;
      }
    } else if (comp.type === 'TABLE') {
      previewMarkup = `<div class="flex items-center gap-1.5 text-muted-foreground"><i data-lucide="table" class="w-6 h-6 text-sky-400"></i> <span class="text-[11px] font-mono">Table Matrix (${comp.table_info ? comp.table_info.rows + 'x' + comp.table_info.cols : 'Data Grid'})</span></div>`;
    } else {
      previewMarkup = `<div class="truncate max-w-full text-center text-xs" style="${inlineStyle}">${escapeHtml(sample)}</div>`;
    }

    card.innerHTML = `
      <div class="flex items-center justify-between">
        <span class="component-type-badge">${comp.type || 'SHAPE'}</span>
        <span class="text-[10px] font-mono text-muted-foreground truncate max-w-[120px]">${comp.source_file || ''}</span>
      </div>

      <div class="component-card-preview">
        ${previewMarkup}
      </div>

      <div class="space-y-1">
        <div class="font-semibold text-xs text-foreground truncate">${escapeHtml(comp.label || 'Component')}</div>
        <div class="text-[11px] text-muted-foreground line-clamp-2 leading-relaxed">${escapeHtml(comp.description || 'Template element')}</div>
      </div>
    `;

    grid.appendChild(card);
  });

  refreshIcons();
}

function openComponentModal(comp) {
  currentModalComponent = comp;
  const modal = document.getElementById('comp-detail-modal');
  if (!modal) return;

  document.getElementById('comp-modal-label').innerText = comp.label || 'Component Schema';
  document.getElementById('comp-modal-desc').innerText = comp.description || '';
  document.getElementById('comp-modal-source').innerText = `Source: ${comp.source_file} (Slide ${comp.slide_index + 1})`;

  const preview = document.getElementById('comp-modal-preview');
  if (comp.type === 'PICTURE' && comp.image_path) {
    const imgName = comp.image_path.split(/[\\/]/).pop();
    preview.innerHTML = `<img src="/api/components/image/${encodeURIComponent(imgName)}" alt="${escapeHtml(comp.label)}" class="max-h-36 max-w-full object-contain rounded">`;
  } else if (comp.sample_text) {
    preview.innerHTML = `<div class="text-sm font-medium text-foreground text-center">${escapeHtml(comp.sample_text)}</div>`;
  } else {
    preview.innerHTML = `<div class="text-xs text-muted-foreground font-mono">[${comp.type}] ${comp.label}</div>`;
  }

  document.getElementById('comp-modal-json').innerText = JSON.stringify(comp, null, 2);
  modal.classList.remove('hidden');
  refreshIcons();
}

function closeComponentModal() {
  const modal = document.getElementById('comp-detail-modal');
  if (modal) modal.classList.add('hidden');
}

function copyModalComponentJson() {
  if (!currentModalComponent) return;
  navigator.clipboard.writeText(JSON.stringify(currentModalComponent, null, 2)).then(() => {
    showToast('Component JSON schema copied', 'success', 2000);
  });
}

async function runComponentsExtraction() {
  const viewer = document.getElementById('components-json-viewer');
  viewer.value = '[*] Scanning data/*.pptx templates and extracting primitives...';
  showToast('Extracting template components...', 'info');

  try {
    const res = await fetch('/api/components/extract', { method: 'POST' });
    const data = await res.json();
    if (data.success) {
      document.getElementById('comp-count-badge').innerText = `Components: ${data.count}`;
      viewer.value = JSON.stringify(data.catalog, null, 2);
      loadComponentsCatalog();
      showToast(`Extraction complete. Discovered ${data.count} components.`, 'success');
    }
  } catch (err) {
    showToast(`Extraction error: ${err.message}`, 'error');
  }
}

// -------------------------------------------------------------
// 5. AUTONOMOUS AI AGENT (Shadcn Chat & Google Gemini Input Area)
// -------------------------------------------------------------
let enableSearchTools = true;
let enablePptxTools = true;
let isTemplateNotesAttached = false;
let attachedNotesCache = '';

function toggleSearchTools() {
  enableSearchTools = !enableSearchTools;
  const btn = document.getElementById('tool-btn-globe');
  if (enableSearchTools) {
    btn.className = 'gemini-icon-btn tool-btn-active text-sky-400';
    btn.title = 'Web Search & Fetch Tools: ENABLED';
    showToast('Web Search & Fetch tools enabled for AI Agent', 'info', 2000);
  } else {
    btn.className = 'gemini-icon-btn text-muted-foreground opacity-40';
    btn.title = 'Web Search & Fetch Tools: DISABLED';
    showToast('Web Search & Fetch tools disabled', 'warning', 2000);
  }
}

function togglePptxTools() {
  enablePptxTools = !enablePptxTools;
  const btn = document.getElementById('tool-btn-pptx');
  if (enablePptxTools) {
    btn.className = 'gemini-icon-btn tool-btn-active text-primary';
    btn.title = 'Slide Synthesis Tools: ENABLED (Generate, Edit & Heal PPTX)';
    showToast('Slide Synthesis & Editing tools enabled for AI Agent', 'info', 2000);
  } else {
    btn.className = 'gemini-icon-btn text-muted-foreground opacity-40';
    btn.title = 'Slide Synthesis Tools: DISABLED';
    showToast('Slide Synthesis & Editing tools disabled', 'warning', 2000);
  }
}

async function toggleAttachTemplateNotes() {
  isTemplateNotesAttached = !isTemplateNotesAttached;
  const banner = document.getElementById('attached-note-banner');
  const btn = document.getElementById('tool-btn-inspect');

  if (isTemplateNotesAttached) {
    try {
      const res = await fetch('/api/templates/notes');
      const data = await res.json();
      attachedNotesCache = data.content || '';
      banner.classList.remove('hidden');
      btn.className = 'gemini-icon-btn tool-btn-active text-amber-400';
      showToast('Attached data/NOTE.md design context to prompt', 'success', 2500);
    } catch (e) {
      showToast('Could not load NOTE.md notes', 'error');
    }
  } else {
    banner.classList.add('hidden');
    btn.className = 'gemini-icon-btn text-muted-foreground';
    attachedNotesCache = '';
    showToast('Detached template notes', 'info', 2000);
  }
}

function detachNoteFromPrompt() {
  isTemplateNotesAttached = false;
  const banner = document.getElementById('attached-note-banner');
  const btn = document.getElementById('tool-btn-inspect');
  if (banner) banner.classList.add('hidden');
  if (btn) btn.className = 'gemini-icon-btn text-muted-foreground';
  attachedNotesCache = '';
}

function autoResizeGeminiInput(textarea) {
  textarea.style.height = 'auto';
  textarea.style.height = Math.min(textarea.scrollHeight, 140) + 'px';
}

function handleGeminiInputKeydown(event) {
  if (event.key === 'Enter' && !event.shiftKey) {
    event.preventDefault();
    sendAgentPrompt();
  }
}

// -------------------------------------------------------------
// CHAT TRANSCRIPT & MESSAGE ACTIONS (Copy / Edit / Export)
// -------------------------------------------------------------
let chatHistoryRecords = [];

function recordChatMessage(role, text, timeStr) {
  chatHistoryRecords.push({
    id: 'msg_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
    role: role,
    text: text,
    time: timeStr || new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
  });
}

function copyMessageText(btn) {
  const item = btn.closest('.chat-message-item');
  if (!item) return;
  const bubble = item.querySelector('.chat-bubble');
  if (!bubble) return;

  const rawText = item.getAttribute('data-raw-content') || bubble.innerText;
  navigator.clipboard.writeText(rawText).then(() => {
    const origHtml = btn.innerHTML;
    btn.innerHTML = `<i data-lucide="check" class="w-3 h-3 text-emerald-400"></i> Copied`;
    refreshIcons();
    setTimeout(() => {
      btn.innerHTML = origHtml;
      refreshIcons();
    }, 2000);
    showToast('Message copied to clipboard', 'success', 2000);
  }).catch(() => {
    showToast('Failed to copy message', 'error');
  });
}

function editUserPrompt(btn) {
  const item = btn.closest('.chat-message-item');
  if (!item) return;
  const rawText = item.getAttribute('data-raw-content') || '';
  const input = document.getElementById('agent-prompt-input');
  if (input) {
    input.value = rawText;
    autoResizeGeminiInput(input);
    input.focus();
    showToast('Prompt loaded into input for editing', 'info', 2000);
  }
}

function exportChatMarkdown() {
  if (chatHistoryRecords.length === 0) {
    showToast('No messages in chat history to export', 'warning');
    return;
  }

  let mdContent = `# PrismPresenter — AI Agent Conversation Transcript\n\n`;
  mdContent += `*Exported on ${new Date().toLocaleString()}*\n\n---\n\n`;

  chatHistoryRecords.forEach(msg => {
    const sender = msg.role === 'user' ? '👤 User' : '🤖 PrismPresenter AI Agent';
    mdContent += `### ${sender} (${msg.time})\n\n${msg.text}\n\n---\n\n`;
  });

  const blob = new Blob([mdContent], { type: 'text/markdown;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `prismpresenter-chat-${Date.now()}.md`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('Exported chat transcript as Markdown (.md)', 'success');
}

function exportChatJson() {
  if (chatHistoryRecords.length === 0) {
    showToast('No messages in chat history to export', 'warning');
    return;
  }

  const exportData = {
    app: "PrismPresenter",
    version: "v0.3",
    exported_at: new Date().toISOString(),
    messages: chatHistoryRecords
  };

  const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `prismpresenter-chat-${Date.now()}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('Exported chat transcript as JSON (.json)', 'success');
}

function copyFullChatTranscript() {
  if (chatHistoryRecords.length === 0) {
    showToast('No messages to copy', 'warning');
    return;
  }

  let transcript = `PRISMPRESENTER AI TERMINAL TRANSCRIPT (${new Date().toLocaleString()})\n\n`;
  chatHistoryRecords.forEach(msg => {
    const sender = msg.role === 'user' ? 'USER' : 'AGENT';
    transcript += `[${msg.time}] ${sender}:\n${msg.text}\n\n`;
  });

  navigator.clipboard.writeText(transcript).then(() => {
    showToast('Full chat transcript copied to clipboard', 'success');
  }).catch(() => {
    showToast('Failed to copy transcript', 'error');
  });
}

function clearAgentChat() {
  chatHistoryRecords = [];
  const chatMessages = document.getElementById('agent-chat-messages');
  chatMessages.innerHTML = `
    <div class="chat-message-item assistant" data-raw-content="Chat cleared. Ready for your presentation instructions.">
      <div class="chat-avatar ai-avatar">
        <i data-lucide="sparkles" class="w-4 h-4"></i>
      </div>
      <div class="chat-bubble-content">
        <div class="chat-bubble">
          Chat cleared. Ready for your presentation instructions.
        </div>
        <div class="chat-meta">
          <div class="chat-meta-info">
            <span>PrismPresenter Agent</span> • <span>System</span>
          </div>
          <div class="chat-message-actions">
            <button class="chat-action-btn" onclick="copyMessageText(this)" title="Copy text">
              <i data-lucide="copy" class="w-3 h-3"></i>
            </button>
          </div>
        </div>
      </div>
    </div>
  `;
  refreshIcons();
}

function applyPromptChip(text) {
  const input = document.getElementById('agent-prompt-input');
  if (input) {
    input.value = text;
    autoResizeGeminiInput(input);
    input.focus();
  }
}

function appendUserChatMessage(text, hasAttachedNotes = false) {
  const container = document.getElementById('agent-chat-messages');
  const item = document.createElement('div');
  item.className = 'chat-message-item user';
  item.setAttribute('data-raw-content', text);

  const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  recordChatMessage('user', text, timeStr);

  const badgeHtml = hasAttachedNotes
    ? `<div class="text-[10px] font-mono bg-black/30 border border-white/20 px-2 py-0.5 rounded-md mb-1 flex items-center gap-1"><i data-lucide="paperclip" class="w-3 h-3 text-amber-300"></i> Context: data/NOTE.md attached</div>`
    : '';

  item.innerHTML = `
    <div class="chat-avatar user-avatar">
      <i data-lucide="user" class="w-4 h-4"></i>
    </div>
    <div class="chat-bubble-content">
      <div class="chat-bubble font-sans">
        ${badgeHtml}
        ${escapeHtml(text)}
      </div>
      <div class="chat-meta">
        <div class="chat-meta-info">
          <span>You</span> • <span>${timeStr}</span>
        </div>
        <div class="chat-message-actions">
          <button class="chat-action-btn" onclick="copyMessageText(this)" title="Copy prompt">
            <i data-lucide="copy" class="w-3 h-3"></i> Copy
          </button>
          <button class="chat-action-btn" onclick="editUserPrompt(this)" title="Edit and re-use prompt">
            <i data-lucide="pencil" class="w-3 h-3"></i> Edit
          </button>
        </div>
      </div>
    </div>
  `;

  container.appendChild(item);
  container.scrollTop = container.scrollHeight;
  refreshIcons();
}

function appendAssistantChatContainer() {
  const container = document.getElementById('agent-chat-messages');
  const item = document.createElement('div');
  item.className = 'chat-message-item assistant';

  const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  item.innerHTML = `
    <div class="chat-avatar ai-avatar">
      <i data-lucide="sparkles" class="w-4 h-4"></i>
    </div>
    <div class="chat-bubble-content w-full">
      <div class="chat-bubble font-sans text-xs leading-relaxed" id="ai-active-reply">
        <div class="flex items-center gap-2 text-muted-foreground text-xs">
          <span class="w-1.5 h-1.5 rounded-full bg-primary animate-pulse"></span>
          Thinking and orchestrating tools...
        </div>
      </div>
      <div class="chat-reasoning-accordion hidden" id="ai-active-reasoning">
        <div class="chat-reasoning-summary" onclick="toggleReasoningAccordion(this)">
          <i data-lucide="chevron-right" class="w-3.5 h-3.5 transition-transform"></i>
          <span>Execution Details & Reasoning Trace</span>
        </div>
        <div class="chat-reasoning-logs"></div>
      </div>
      <div class="chat-meta">
        <div class="chat-meta-info">
          <span>PrismPresenter Agent</span> • <span>${timeStr}</span>
        </div>
        <div class="chat-message-actions">
          <button class="chat-action-btn" onclick="copyMessageText(this)" title="Copy response">
            <i data-lucide="copy" class="w-3 h-3"></i> Copy
          </button>
        </div>
      </div>
    </div>
  `;

  container.appendChild(item);
  container.scrollTop = container.scrollHeight;
  refreshIcons();

  return item;
}

function toggleReasoningAccordion(el) {
  const logs = el.nextElementSibling;
  const icon = el.querySelector('i');
  if (logs.classList.contains('hidden')) {
    logs.classList.remove('hidden');
    if (icon) icon.style.transform = 'rotate(90deg)';
  } else {
    logs.classList.add('hidden');
    if (icon) icon.style.transform = 'rotate(0deg)';
  }
}

function renderMarkdownContent(rawText) {
  if (window.marked) {
    try {
      marked.setOptions({
        breaks: true,
        gfm: true
      });
      let html = marked.parse(rawText);

      // Post-process pre/code blocks with copy snippet headers
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');

      doc.querySelectorAll('pre').forEach(pre => {
        const codeElem = pre.querySelector('code');
        const codeText = codeElem ? codeElem.innerText : pre.innerText;
        let lang = 'code';
        if (codeElem && codeElem.className) {
          const m = codeElem.className.match(/language-(\w+)/);
          if (m) lang = m[1];
        }

        const wrapper = doc.createElement('div');
        wrapper.className = 'code-block-wrapper';
        wrapper.innerHTML = `
          <div class="code-block-header">
            <span>${lang}</span>
            <button type="button" class="code-block-copy-btn" onclick="copySnippetCode(this)">
              <i data-lucide="copy" class="w-3 h-3"></i> Copy
            </button>
          </div>
          <pre class="code-block-content"><code>${escapeHtml(codeText)}</code></pre>
        `;

        pre.parentNode.replaceChild(wrapper, pre);
      });

      return doc.body.innerHTML;
    } catch (e) {
      console.error('Markdown parse error:', e);
    }
  }
  return `<div class="whitespace-pre-wrap">${escapeHtml(rawText)}</div>`;
}

function copySnippetCode(btn) {
  const wrapper = btn.closest('.code-block-wrapper');
  if (!wrapper) return;
  const code = wrapper.querySelector('code');
  if (code) {
    navigator.clipboard.writeText(code.innerText).then(() => {
      const origText = btn.innerHTML;
      btn.innerHTML = `<i data-lucide="check" class="w-3 h-3 text-emerald-400"></i> Copied!`;
      refreshIcons();
      setTimeout(() => {
        btn.innerHTML = origText;
        refreshIcons();
      }, 2000);
      showToast('Code snippet copied to clipboard', 'success', 2000);
    }).catch(() => {
      showToast('Failed to copy snippet', 'error');
    });
  }
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.innerText = text;
  return div.innerHTML;
}

async function sendAgentPrompt() {
  const input = document.getElementById('agent-prompt-input');
  let prompt = input.value.trim();
  if (!prompt) return;

  const hasAttached = isTemplateNotesAttached;
  if (isTemplateNotesAttached && attachedNotesCache) {
    prompt = `[ATTACHED REFERENCE CONTEXT: data/NOTE.md]\n\`\`\`markdown\n${attachedNotesCache}\n\`\`\`\n\n[USER INSTRUCTION]:\n${prompt}`;
  }

  input.value = '';
  autoResizeGeminiInput(input);
  if (isTemplateNotesAttached) {
    detachNoteFromPrompt();
  }

  const btn = document.getElementById('btn-agent-send');
  const indicator = document.getElementById('agent-typing-indicator');

  btn.disabled = true;
  if (indicator) indicator.classList.remove('hidden');
  setSystemStatus('REASONING...', true);

  // Render User Chat Bubble
  const displayPrompt = hasAttached ? prompt.split('[USER INSTRUCTION]:\n').pop() : prompt;
  appendUserChatMessage(displayPrompt, hasAttached);

  // Render AI Response Placeholder Bubble
  const aiMsgElem = appendAssistantChatContainer();
  const replyBubble = aiMsgElem.querySelector('#ai-active-reply');
  const reasoningAccordion = aiMsgElem.querySelector('#ai-active-reasoning');
  const reasoningLogs = reasoningAccordion.querySelector('.chat-reasoning-logs');

  replyBubble.removeAttribute('id');
  reasoningAccordion.removeAttribute('id');

  try {
    const res = await fetch('/api/agent/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        prompt: prompt,
        enable_search: enableSearchTools,
        enable_pptx_tools: enablePptxTools
      })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const evtSource = new EventSource(`/api/generator/stream/${data.job_id}`);

    evtSource.addEventListener('log', (e) => {
      const d = JSON.parse(e.data);
      reasoningAccordion.classList.remove('hidden');
      const logLine = document.createElement('div');
      logLine.className = 'text-[11px] font-mono py-0.5 border-b border-border/20 text-muted-foreground';
      logLine.innerText = `[${d.time}] ${d.message}`;
      reasoningLogs.appendChild(logLine);
      reasoningLogs.scrollTop = reasoningLogs.scrollHeight;
    });

    evtSource.addEventListener('completed', (e) => {
      const d = JSON.parse(e.data);
      aiMsgElem.setAttribute('data-raw-content', d.response);
      recordChatMessage('assistant', d.response, new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
      replyBubble.innerHTML = `<div class="chat-markdown">${renderMarkdownContent(d.response)}</div>`;
      setSystemStatus('READY');
      btn.disabled = false;
      if (indicator) indicator.classList.add('hidden');
      refreshIcons();
      showToast('Agent reasoning task complete', 'success');
    });

    evtSource.addEventListener('error', () => {
      replyBubble.innerHTML = `<span class="text-destructive font-medium">Agent execution encountered an error.</span>`;
      setSystemStatus('ERROR');
      btn.disabled = false;
      if (indicator) indicator.classList.add('hidden');
      showToast('Agent execution error', 'error');
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    replyBubble.innerHTML = `<span class="text-destructive font-medium">Agent failed: ${escapeHtml(err.message)}</span>`;
    setSystemStatus('ERROR');
    btn.disabled = false;
    if (indicator) indicator.classList.add('hidden');
    showToast(`Agent error: ${err.message}`, 'error');
  }
}

// -------------------------------------------------------------
// 6. SETTINGS & CONFIGURATION
// -------------------------------------------------------------
async function loadConfigSettings() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    if (data.success && data.config) {
      const cfg = data.config;
      document.getElementById('cfg-url').value = cfg.NINEROUTER_URL || '';
      document.getElementById('cfg-key').value = cfg.NINEROUTER_KEY || '';
      document.getElementById('cfg-chat-model').value = cfg.NINEROUTER_CHAT_MODEL || '';
      document.getElementById('cfg-search-model').value = cfg.NINEROUTER_SEARCH_MODEL || '';
      document.getElementById('cfg-fetch-model').value = cfg.NINEROUTER_FETCH_MODEL || '';
      document.getElementById('cfg-image-model').value = cfg.NINEROUTER_IMAGE_MODEL || '';
      const timeoutEl = document.getElementById('cfg-timeout');
      if (timeoutEl) timeoutEl.value = cfg.LLM_TIMEOUT || 300;

      applyConfigAndMetadata(data);
    }
  } catch (err) {
    console.error('Failed to load settings', err);
  }
}

async function saveConfigSettings() {
  const timeoutVal = parseInt(document.getElementById('cfg-timeout')?.value.trim() || '300', 10);
  const config = {
    NINEROUTER_URL: document.getElementById('cfg-url').value.trim(),
    NINEROUTER_KEY: document.getElementById('cfg-key').value.trim(),
    NINEROUTER_CHAT_MODEL: document.getElementById('cfg-chat-model').value.trim(),
    NINEROUTER_SEARCH_MODEL: document.getElementById('cfg-search-model').value.trim(),
    NINEROUTER_FETCH_MODEL: document.getElementById('cfg-fetch-model').value.trim(),
    NINEROUTER_IMAGE_MODEL: document.getElementById('cfg-image-model').value.trim(),
    LLM_TIMEOUT: isNaN(timeoutVal) ? 300 : timeoutVal,
    PURE_PIL_ACTIVE: true
  };

  try {
    const res = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ config })
    });
    const data = await res.json();
    if (data.success) {
      showToast('Configuration saved to .env', 'success');
      loadConfigBadge();
    }
  } catch (err) {
    showToast(`Save failed: ${err.message}`, 'error');
  }
}
