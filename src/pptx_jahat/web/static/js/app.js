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
let genDiagnosticsSteps = [];

// Human Touch Workflow State
let currentHumanTouchJobId = null;
let currentHumanTouchStep = null;
let currentHumanTouchData = null;
let currentHumanTouchView = 'visual';

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
    loadGeneratorStructures();
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
    loadGeneratorDiagnostics();
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
  if (badge) {
    badge.innerText = `MODEL: ${modelName}`;
    badge.title = `Active 9Router Model: ${cfg.NINEROUTER_CHAT_MODEL || ''}`;
  }

  const geminiModel = document.getElementById('gemini-model-name');
  if (geminiModel) geminiModel.innerText = modelName;

  const chatInput = document.getElementById('cfg-chat-model');
  if (chatInput && cfg.NINEROUTER_CHAT_MODEL && chatInput !== document.activeElement) {
    chatInput.value = cfg.NINEROUTER_CHAT_MODEL;
  }

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

  // Update agent status badges with effective model and thinking levels
  updateAllAgentStatusBadges(data);
}

async function refreshModelMetadataUI() {
  showToast('Querying 9Router model capabilities...', 'info', 1500);
  await loadConfigBadge();
  await loadConfigSettings();
  showToast('Updated model intelligence from 9Router', 'success', 2000);
}

// -------------------------------------------------------------
// 1. SLIDE GENERATOR & PIPELINE DIAGNOSTICS
// -------------------------------------------------------------
function getDefaultDiagnosticsSteps() {
  return [
    {
      id: "step_1",
      name: "Step 1: Scan & Inspect Templates",
      status: "pending",
      input: "Template catalog (data/), selected template style, visual slide screenshots",
      output: "Awaiting template inspection...",
      duration: null
    },
    {
      id: "step_1_5",
      name: "Step 1.5: Storyboard Schema Resolution",
      status: "pending",
      input: "Storyboard schema (.md) for detection, restructuring, and layout blueprints",
      output: "Awaiting schema resolution...",
      duration: null
    },
    {
      id: "step_2",
      name: "Step 2: Ingest & Parse Source Content",
      status: "pending",
      input: "Source documents (Word .docx, PPTX, MD, TXT, Images, Audio) and direct notes",
      output: "Awaiting content ingestion...",
      duration: null
    },
    {
      id: "step_2_5",
      name: "Step 2.5: Storyboard Restructuring (AI Agent)",
      status: "pending",
      input: "Extracted document sections & active restructure schema rules",
      output: "Awaiting restructure agent...",
      duration: null
    },
    {
      id: "step_3",
      name: "Step 3: Vision AI Reasoning & Slide Selection",
      status: "pending",
      input: "Template slide screenshots, content sections, schema archetype rules",
      output: "Awaiting AI reasoning and slide selection...",
      duration: null
    },
    {
      id: "step_4",
      name: "Step 4: Deck Assembly & Slide Cloning",
      status: "pending",
      input: "AI synthesis plan, source template slides, target layout parameters",
      output: "Awaiting presentation assembly...",
      duration: null
    },
    {
      id: "step_5",
      name: "Step 5: SlideCheck QA & Integrity Verification",
      status: "pending",
      input: "Assembled PPTX presentation, template typography, geometry boundaries",
      output: "Awaiting SlideCheck QA & font auto-healing...",
      duration: null
    }
  ];
}

async function loadGeneratorDiagnostics(jobId = null) {
  try {
    const url = jobId ? `/api/generator/diagnostics?job_id=${encodeURIComponent(jobId)}` : '/api/generator/diagnostics';
    const res = await fetch(url);
    const data = await res.json();
    if (data.success && Array.isArray(data.diagnostics) && data.diagnostics.length > 0) {
      genDiagnosticsSteps = data.diagnostics;
    } else if (genDiagnosticsSteps.length === 0) {
      genDiagnosticsSteps = getDefaultDiagnosticsSteps();
    }
  } catch (e) {
    if (genDiagnosticsSteps.length === 0) {
      genDiagnosticsSteps = getDefaultDiagnosticsSteps();
    }
  }
  renderGenDiagnostics();
}

function refreshDiagnostics() {
  loadGeneratorDiagnostics();
  showToast('Diagnostics refreshed', 'info');
}

function resetGenDiagnostics() {
  genDiagnosticsSteps = getDefaultDiagnosticsSteps();
  renderGenDiagnostics();
}

function setGenDiagnostics(steps) {
  if (Array.isArray(steps) && steps.length > 0) {
    genDiagnosticsSteps = steps;
    renderGenDiagnostics();
  }
}

function updateGenStep(stepData) {
  if (!stepData || !stepData.id) return;
  const idx = genDiagnosticsSteps.findIndex(s => s.id === stepData.id);
  if (idx !== -1) {
    genDiagnosticsSteps[idx] = { ...genDiagnosticsSteps[idx], ...stepData };
  } else {
    genDiagnosticsSteps.push(stepData);
  }
  renderGenDiagnostics();
}

function markActiveGenStepFailed(errorMsg) {
  const activeStep = genDiagnosticsSteps.find(s => s.status === 'running');
  if (activeStep) {
    activeStep.status = 'failed';
    activeStep.output = `Error: ${errorMsg || 'Pipeline terminated with error'}`;
  }
  renderGenDiagnostics();
}

function getStatusBadgeHtml(status, duration) {
  const durStr = duration ? `<span class="text-[10px] opacity-80 font-normal ml-0.5">(${escapeHtml(duration)})</span>` : '';
  switch (status) {
    case 'completed':
      return `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-medium bg-emerald-500/15 text-emerald-400 border border-emerald-500/30 shadow-sm"><i data-lucide="check-circle-2" class="w-3.5 h-3.5 text-emerald-400"></i> Completed ${durStr}</span>`;
    case 'running':
      return `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-medium bg-amber-500/15 text-amber-400 border border-amber-500/30 animate-pulse shadow-sm"><i data-lucide="loader-2" class="w-3.5 h-3.5 animate-spin text-amber-400"></i> Running...</span>`;
    case 'skipped':
      return `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-medium bg-secondary text-muted-foreground border border-border"><i data-lucide="fast-forward" class="w-3.5 h-3.5"></i> Skipped ${durStr}</span>`;
    case 'failed':
      return `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-medium bg-destructive/20 text-destructive border border-destructive/40 shadow-sm"><i data-lucide="alert-triangle" class="w-3.5 h-3.5 text-destructive"></i> Failed ${durStr}</span>`;
    case 'pending':
    default:
      return `<span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-mono font-medium bg-secondary/50 text-muted-foreground border border-border/60"><i data-lucide="clock" class="w-3.5 h-3.5"></i> Pending</span>`;
  }
}

function renderDiagnosticsSummaryPills(containerId) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const total = genDiagnosticsSteps.length;
  const completed = genDiagnosticsSteps.filter(s => s.status === 'completed').length;
  const running = genDiagnosticsSteps.filter(s => s.status === 'running').length;
  const failed = genDiagnosticsSteps.filter(s => s.status === 'failed').length;
  const skipped = genDiagnosticsSteps.filter(s => s.status === 'skipped').length;
  const pending = genDiagnosticsSteps.filter(s => s.status === 'pending').length;

  container.innerHTML = `
    <span class="px-2 py-0.5 rounded bg-secondary text-foreground text-[10px] border border-border">Total: ${total}</span>
    <span class="px-2 py-0.5 rounded bg-emerald-500/15 text-emerald-400 text-[10px] border border-emerald-500/30 font-bold">✓ ${completed} Done</span>
    ${running > 0 ? `<span class="px-2 py-0.5 rounded bg-amber-500/15 text-amber-400 text-[10px] border border-amber-500/30 font-bold animate-pulse">⚡ ${running} Active</span>` : ''}
    ${failed > 0 ? `<span class="px-2 py-0.5 rounded bg-destructive/20 text-destructive text-[10px] border border-destructive/40 font-bold">✗ ${failed} Failed</span>` : ''}
    ${skipped > 0 ? `<span class="px-2 py-0.5 rounded bg-muted/40 text-muted-foreground text-[10px] border border-border/50">⏭ ${skipped} Skipped</span>` : ''}
    ${pending > 0 ? `<span class="px-2 py-0.5 rounded bg-secondary/60 text-muted-foreground text-[10px] border border-border/50">⏳ ${pending} Queued</span>` : ''}
  `;
}

function renderStepCardHtml(step, index) {
  const isRunning = step.status === 'running';
  const isCompleted = step.status === 'completed';
  const isFailed = step.status === 'failed';

  const borderClass = isRunning
    ? 'border-amber-500/60 bg-amber-950/10 shadow-md shadow-amber-500/5 ring-1 ring-amber-500/30'
    : isCompleted
      ? 'border-emerald-500/30 bg-card/50'
      : isFailed
        ? 'border-destructive/60 bg-destructive/5'
        : 'border-border/60 bg-card/30';

  return `
    <div class="rounded-lg border ${borderClass} p-3.5 transition duration-200">
      <div class="flex items-center justify-between flex-wrap gap-2 pb-2.5 border-b border-border/40">
        <div class="flex items-center gap-2">
          <span class="w-6 h-6 rounded-md bg-secondary border border-border/80 flex items-center justify-center font-mono text-[11px] font-bold text-foreground">
            ${index + 1}
          </span>
          <span class="font-semibold text-xs text-foreground tracking-tight">${escapeHtml(step.name || '')}</span>
        </div>
        <div>
          ${getStatusBadgeHtml(step.status, step.duration)}
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-2.5 mt-2.5 text-xs">
        <!-- INPUT BLOCK -->
        <div class="rounded-md border border-border/50 bg-background/80 p-2.5 flex flex-col">
          <div class="flex items-center gap-1.5 text-[11px] font-semibold text-primary mb-1">
            <i data-lucide="arrow-down-right" class="w-3.5 h-3.5 text-primary"></i>
            <span>Input</span>
          </div>
          <div class="text-[11px] font-mono text-muted-foreground/90 whitespace-pre-wrap leading-relaxed flex-1 overflow-x-auto select-text">
            ${escapeHtml(step.input || 'No input defined.')}
          </div>
        </div>

        <!-- OUTPUT BLOCK -->
        <div class="rounded-md border border-border/50 bg-background/80 p-2.5 flex flex-col">
          <div class="flex items-center gap-1.5 text-[11px] font-semibold ${isCompleted ? 'text-emerald-400' : isFailed ? 'text-destructive' : 'text-cyan-400'} mb-1">
            <i data-lucide="arrow-up-right" class="w-3.5 h-3.5"></i>
            <span>Output</span>
          </div>
          <div class="text-[11px] font-mono ${isCompleted ? 'text-foreground/90' : isFailed ? 'text-destructive/90 font-bold' : 'text-muted-foreground/90'} whitespace-pre-wrap leading-relaxed flex-1 overflow-x-auto select-text">
            ${escapeHtml(step.output || 'Pending execution...')}
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderGenDiagnostics() {
  const mainList = document.getElementById('gen-diagnostics-steps-list');
  const subList = document.getElementById('gen-subtab-diagnostics-steps-list');

  if (genDiagnosticsSteps.length === 0) {
    genDiagnosticsSteps = getDefaultDiagnosticsSteps();
  }

  const cardsHtml = genDiagnosticsSteps.map((step, idx) => renderStepCardHtml(step, idx)).join('');

  if (mainList) {
    mainList.innerHTML = cardsHtml;
  }
  if (subList) {
    subList.innerHTML = cardsHtml;
  }

  renderDiagnosticsSummaryPills('diag-summary-pills');
  renderDiagnosticsSummaryPills('gen-subtab-diag-summary-pills');

  if (window.lucide) {
    lucide.createIcons();
  }
}

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
    checkGeneratorTemplateFonts();
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

function loadStructuresList() {
  return loadGeneratorStructures();
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
  const enableVerification = document.getElementById('gen-enable-verification')?.checked ?? true;
  const verificationRounds = parseInt(document.getElementById('gen-verification-rounds')?.value || '3', 10);
  const enableHumanTouch = document.getElementById('gen-enable-human-touch')?.checked ?? false;
  const htExtract = document.getElementById('gen-ht-extract')?.checked ?? true;
  const htRestructure = document.getElementById('gen-ht-restructure')?.checked ?? true;
  const htVerify = document.getElementById('gen-ht-verify')?.checked ?? true;
  const htAfterDone = document.getElementById('gen-ht-after-done')?.checked ?? true;

  const humanTouchSteps = [];
  if (htExtract) humanTouchSteps.push('extract');
  if (htRestructure) humanTouchSteps.push('restructure');
  if (htVerify) humanTouchSteps.push('verify');
  if (htAfterDone) humanTouchSteps.push('after_done');

  resetGenDiagnostics();

  try {
    if (enableVerification) {
      appendGenLog(`[*] Visual Template Verification Agent ACTIVE: Auditing slides up to ${verificationRounds} iterative round(s).`);
    } else {
      appendGenLog('[*] Visual Template Verification Agent DISABLED by user at generation time.');
    }

    const activeChatModel = document.getElementById('cfg-chat-model')?.value.trim() || undefined;
    if (activeChatModel) {
      appendGenLog(`[*] Active 9Router AI Model: '${activeChatModel}'`);
    }

    if (enableHumanTouch) {
      appendGenLog(`[*] Human Touch Workflow ACTIVE: Will pause at steps [${humanTouchSteps.join(', ')}] for human review.`);
    }

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
        enable_verification: enableVerification,
        verification_rounds: verificationRounds,
        chat_model: activeChatModel,
        enable_human_touch: enableHumanTouch,
        human_touch_steps: humanTouchSteps,
        output_path: outputPath,
        timeout: timeoutVal,
        font_fallbacks: currentConfiguredFontFallbacks
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

  evtSource.addEventListener('step_update', (e) => {
    try {
      const d = JSON.parse(e.data);
      updateGenStep(d);
    } catch (_) {}
  });

  evtSource.addEventListener('ai_images', (e) => {
    const d = JSON.parse(e.data);
    aiTestImages = d.images || [];
    aiTestIdx = 0;
    updateAiTestDisplay();
  });

  evtSource.addEventListener('human_review', (e) => {
    try {
      const d = JSON.parse(e.data);
      if (d.step === 'font_fallback') {
        openFontFallbackModal(d);
      } else {
        onHumanReviewEvent(d);
      }
    } catch (err) {
      console.error('Error handling human_review event:', err);
    }
  });

  evtSource.addEventListener('human_action_accepted', (e) => {
    closeHumanTouchModal();
    closeFontFallbackModal();
  });

  evtSource.addEventListener('human_review_error', (e) => {
    try {
      const d = JSON.parse(e.data);
      showToast(`Human Touch notice: ${d.error}`, 'error');
      const statusSpan = document.getElementById('ht-rerun-status');
      if (statusSpan) statusSpan.innerText = `Error: ${d.error}`;
      const rerunBtn = document.getElementById('btn-ht-rerun');
      if (rerunBtn) rerunBtn.disabled = false;
    } catch (_) {}
  });

  evtSource.addEventListener('completed', (e) => {
    const d = JSON.parse(e.data);
    currentGeneratedPptx = d.pptx_path;
    genSlides = d.previews || [];
    genSlideIdx = 0;
    visualSlides = [...genSlides];
    visualSlideIdx = 0;

    if (d.diagnostics && Array.isArray(d.diagnostics)) {
      setGenDiagnostics(d.diagnostics);
    }

    updateGenSlideDisplay(d.engine_name);
    updateVisualSlideDisplay();
    renderSlideThumbnails();

    btnOpen.disabled = false;
    btnDownload.disabled = false;
    btnGen.disabled = false;

    setSystemStatus('READY');
    appendGenLog(`[✓] Completed: ${d.filename}`, 'success');
    showToast(`Generation complete: ${d.filename}`, 'success');

    // Human Touch After-Done highlight
    const enableHumanTouch = document.getElementById('gen-enable-human-touch')?.checked ?? false;
    const htAfterDone = document.getElementById('gen-ht-after-done')?.checked ?? true;
    if (enableHumanTouch && htAfterDone) {
      appendGenLog('[*] Human Touch (After Done): Presentation deck ready. Reply with a custom prompt in the editor below to rerun & modify with AI.');
      document.getElementById('ht-deck-edit-prompt')?.focus();
    }
  });

  evtSource.addEventListener('error', (e) => {
    try {
      const d = JSON.parse(e.data);
      appendGenLog(`[!] Error: ${d.error}`);
      if (d.diagnostics && Array.isArray(d.diagnostics)) {
        setGenDiagnostics(d.diagnostics);
      } else {
        markActiveGenStepFailed(d.error);
      }
    } catch (_) {
      markActiveGenStepFailed('Generation job encountered an unhandled error.');
    }
    setSystemStatus('ERROR');
    btnGen.disabled = false;
    showToast('Slide generation failed.', 'error');
  });

  evtSource.addEventListener('close', () => {
    evtSource.close();
  });
}

let currentGenEngine = '';

function updateGenSlideDisplay(engineName = '') {
  const img = document.getElementById('gen-slide-img');
  const ph = document.getElementById('gen-slide-placeholder');
  const counter = document.getElementById('gen-slide-counter');
  const badge = document.getElementById('gen-engine-badge');
  const prevBtn = document.getElementById('gen-prev-btn');
  const nextBtn = document.getElementById('gen-next-btn');

  if (engineName) {
    currentGenEngine = engineName;
  }

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

  const activeEngine = currentGenEngine || 'Native PowerPoint';
  if (badge) {
    badge.classList.remove('hidden');
    if (activeEngine.includes('PowerPoint')) {
      badge.innerText = 'Native PowerPoint';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 border border-emerald-800 text-emerald-300';
    } else if (activeEngine.includes('Web')) {
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
// 1.5. HUMAN TOUCH WORKFLOW & PRESENTATION ITERATION ENGINE
// -------------------------------------------------------------
function toggleHumanTouchOptions() {
  const checkbox = document.getElementById('gen-enable-human-touch');
  const container = document.getElementById('gen-ht-options-container');
  if (!checkbox || !container) return;
  if (checkbox.checked) {
    container.classList.remove('hidden');
  } else {
    container.classList.add('hidden');
  }
}

function onHumanReviewEvent(payload) {
  currentHumanTouchJobId = payload.job_id;
  currentHumanTouchStep = payload.step;
  currentHumanTouchData = payload.data || {};
  currentHumanTouchView = 'visual';

  const modal = document.getElementById('human-touch-modal');
  const titleEl = document.getElementById('ht-modal-title');
  const badgeEl = document.getElementById('ht-modal-step-badge');
  const subtitleEl = document.getElementById('ht-modal-subtitle');
  const rerunStatus = document.getElementById('ht-rerun-status');
  const promptInput = document.getElementById('ht-reply-prompt-input');

  if (rerunStatus) rerunStatus.innerText = '';
  if (promptInput) promptInput.value = '';

  const stepNames = {
    extract: {
      title: 'Human Touch: Review Extracted Slide Content',
      badge: 'Step 2: Extract Slides',
      subtitle: 'Inspect AI extracted sections, edit text or bullets directly, or reply with instructions to rerun.',
      chips: [
        'Condense into 4 focused slides',
        'Expand with deeper technical explanations',
        'Make bullet points punchier and shorter',
        'Translate slide contents to Persian / فارسی',
        'Focus on key mathematical formulas',
        'Highlight teacher takeaways & examples'
      ]
    },
    restructure: {
      title: 'Human Touch: Review Restructured Storyboard',
      badge: 'Step 2.5: Restructure Slides',
      subtitle: 'Inspect storyboard quadrants, animation steps ①②③, formulas, or reply with instructions to rerun.',
      chips: [
        'Follow 4-quadrant layout: TR -> TL -> BR -> BL',
        'Add numbered circled animation steps (①, ②, ③, ④)',
        'Add teacher callout box: نکته کلیدی',
        'Format math expressions into clean LaTeX',
        'Simplify bullets to 3 per slide',
        'Expand speaker notes for lecture presentation'
      ]
    },
    verify: {
      title: 'Human Touch: Template Alignment & Visual Verification',
      badge: 'Step 4.5: Slide Verification',
      subtitle: 'Inspect chosen template screenshots vs generated slides, converse with the Verification Agent to spot missed elements, and approve healing edits.',
      chips: [
        'Delete unreplaced template formulas',
        'Restore missing subtitle from template',
        'Shape has default placeholder text, fix it',
        'Remove mismatched equation from slide',
        'Fill empty card with topic summary',
        'Make sure header is not overflowing',
        'Don\'t delete bottom badge or icon',
        'Slides look aligned, approve without edits'
      ]
    }
  };

  const info = stepNames[payload.step] || {
    title: `Human Touch: Review ${payload.step}`,
    badge: payload.step,
    subtitle: 'Review AI response, make manual edits, or reply with custom feedback to rerun.',
    chips: ['Refine and optimize for presentation', 'Make text more concise']
  };

  if (titleEl) titleEl.innerText = info.title;
  if (badgeEl) badgeEl.innerText = info.badge;
  if (subtitleEl) subtitleEl.innerText = info.subtitle;

  // Toggle Skip button and Continue button text
  const skipVerifyBtn = document.getElementById('btn-ht-skip-verify');
  const continueBtn = document.getElementById('btn-ht-continue');
  if (payload.step === 'verify') {
    if (skipVerifyBtn) skipVerifyBtn.classList.remove('hidden');
    if (continueBtn) {
      continueBtn.innerHTML = '<i data-lucide="check-check" class="w-3.5 h-3.5"></i> Approve &amp; Apply Healing Edits';
    }
  } else {
    if (skipVerifyBtn) skipVerifyBtn.classList.add('hidden');
    if (continueBtn) {
      continueBtn.innerHTML = '<i data-lucide="check" class="w-3.5 h-3.5"></i> Approve &amp; Continue to Next Step';
    }
  }

  // Render suggestion chips
  const chipsContainer = document.getElementById('ht-suggestion-chips');
  if (chipsContainer) {
    chipsContainer.innerHTML = `
      <span class="text-muted-foreground font-mono self-center">Suggestions:</span>
      ${info.chips.map(c => `
        <button type="button" onclick="setHumanTouchPrompt('${c.replace(/'/g, "\\'")}')"
          class="px-2 py-0.5 rounded bg-secondary/80 hover:bg-secondary text-muted-foreground hover:text-foreground border border-border/60 transition cursor-pointer">
          ${c}
        </button>
      `).join('')}
    `;
  }

  // Set Document Title
  const docTitleInput = document.getElementById('ht-doc-title-input');
  if (docTitleInput) {
    docTitleInput.value = currentHumanTouchData.document_title || 'Presentation';
  }

  // Render Visual View
  if (currentHumanTouchStep === 'verify') {
    renderHumanTouchVerification(currentHumanTouchData);
  } else {
    renderHumanTouchVisual(currentHumanTouchStep, currentHumanTouchData);
  }

  const jsonArea = document.getElementById('ht-raw-json-textarea');
  if (jsonArea) {
    jsonArea.value = JSON.stringify(currentHumanTouchData, null, 2);
  }

  setHumanTouchView('visual');

  if (modal) {
    modal.classList.remove('hidden');
  }
  if (window.lucide) lucide.createIcons();

  showToast(`Human Touch: ${info.badge} ready for review`, 'info');
}

function setHumanTouchPrompt(text) {
  const input = document.getElementById('ht-reply-prompt-input');
  if (input) {
    input.value = text;
    input.focus();
  }
}

function setHumanTouchView(view) {
  currentHumanTouchView = view;
  const visualBtn = document.getElementById('ht-view-visual-btn');
  const jsonBtn = document.getElementById('ht-view-json-btn');
  const visualBox = document.getElementById('ht-visual-container');
  const verifyBox = document.getElementById('ht-verify-container');
  const jsonBox = document.getElementById('ht-json-container');

  if (view === 'visual') {
    // Sync from JSON to Visual if valid
    const jsonArea = document.getElementById('ht-raw-json-textarea');
    if (jsonArea && jsonArea.value.trim()) {
      try {
        const parsed = JSON.parse(jsonArea.value.trim());
        currentHumanTouchData = parsed;
        if (currentHumanTouchStep === 'verify') {
          renderHumanTouchVerification(currentHumanTouchData);
        } else {
          renderHumanTouchVisual(currentHumanTouchStep, currentHumanTouchData);
        }
      } catch (err) {
        showToast('JSON syntax error; preserving current visual fields.', 'warning');
      }
    }
    if (currentHumanTouchStep === 'verify') {
      if (verifyBox) verifyBox.classList.remove('hidden');
      if (visualBox) visualBox.classList.add('hidden');
    } else {
      if (visualBox) visualBox.classList.remove('hidden');
      if (verifyBox) verifyBox.classList.add('hidden');
    }
    if (jsonBox) jsonBox.classList.add('hidden');
    if (visualBtn) {
      visualBtn.className = 'px-2.5 py-1 rounded font-medium bg-background text-foreground shadow-sm';
    }
    if (jsonBtn) {
      jsonBtn.className = 'px-2.5 py-1 rounded font-medium text-muted-foreground hover:text-foreground';
    }
  } else {
    // Sync from Visual to JSON
    currentHumanTouchData = collectHumanTouchDataFromVisual();
    const jsonArea = document.getElementById('ht-raw-json-textarea');
    if (jsonArea) {
      jsonArea.value = JSON.stringify(currentHumanTouchData, null, 2);
    }
    if (visualBox) visualBox.classList.add('hidden');
    if (verifyBox) verifyBox.classList.add('hidden');
    if (jsonBox) jsonBox.classList.remove('hidden');
    if (visualBtn) {
      visualBtn.className = 'px-2.5 py-1 rounded font-medium text-muted-foreground hover:text-foreground';
    }
    if (jsonBtn) {
      jsonBtn.className = 'px-2.5 py-1 rounded font-medium bg-background text-foreground shadow-sm';
    }
  }
}

// -------------------------------------------------------------
// VERIFICATION HUMAN TOUCH FUNCTIONS
// -------------------------------------------------------------
let currentVerifySlideIdx = 0;

function renderHumanTouchVerification(data) {
  if (!data) return;
  const slides = data.slides || [];
  const selector = document.getElementById('ht-verify-slide-selector');
  const overallBadge = document.getElementById('ht-verify-overall-badge');

  if (overallBadge) {
    const totalIssues = data.total_issues_found || 0;
    const totalActions = data.total_actions_planned || 0;
    const roundStr = (data.round && data.max_rounds) ? ` [Round ${data.round}/${data.max_rounds}]` : '';
    if (data.all_correct || (totalIssues === 0 && totalActions === 0)) {
      overallBadge.innerText = `✓ Visual Verification: All Slides Aligned OK${roundStr}`;
      overallBadge.className = 'text-[11px] font-mono px-2.5 py-1 rounded bg-emerald-950/80 border border-emerald-800 text-emerald-300';
    } else {
      overallBadge.innerText = `⚠ Found ${totalIssues} discrepancy item(s) • ${totalActions} healing action(s)${roundStr}`;
      overallBadge.className = 'text-[11px] font-mono px-2.5 py-1 rounded bg-amber-950/80 border border-amber-800 text-amber-300';
    }
  }

  if (selector) {
    selector.innerHTML = '';
    slides.forEach((s, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      const hasIssues = (s.detected_issues && s.detected_issues.length > 0) || (s.edit_structure?.actions?.length > 0);
      const isSelected = (idx === currentVerifySlideIdx);

      btn.className = `px-2.5 py-1 rounded text-xs font-medium border transition cursor-pointer flex items-center gap-1.5 ${
        isSelected
          ? 'bg-primary text-primary-foreground border-primary'
          : hasIssues
            ? 'bg-amber-950/40 text-amber-300 border-amber-800/80 hover:bg-amber-950/70'
            : 'bg-secondary text-muted-foreground hover:text-foreground border-border'
      }`;

      const icon = hasIssues ? '⚠' : '✓';
      btn.innerHTML = `<span>Slide ${s.slide_number}</span><span class="text-[10px] opacity-80">${icon}</span>`;
      btn.onclick = () => selectVerifySlide(idx);
      selector.appendChild(btn);
    });
  }

  if (currentVerifySlideIdx >= slides.length) {
    currentVerifySlideIdx = 0;
  }
  selectVerifySlide(currentVerifySlideIdx);
}

function selectVerifySlide(idx) {
  const slides = currentHumanTouchData.slides || [];
  if (!slides || slides.length === 0) return;
  if (idx < 0 || idx >= slides.length) idx = 0;
  currentVerifySlideIdx = idx;

  // Update button highlights
  const selector = document.getElementById('ht-verify-slide-selector');
  if (selector) {
    Array.from(selector.children).forEach((btn, bIdx) => {
      const s = slides[bIdx];
      const hasIssues = s && ((s.detected_issues && s.detected_issues.length > 0) || (s.edit_structure?.actions?.length > 0));
      const isSelected = (bIdx === currentVerifySlideIdx);
      btn.className = `px-2.5 py-1 rounded text-xs font-medium border transition cursor-pointer flex items-center gap-1.5 ${
        isSelected
          ? 'bg-primary text-primary-foreground border-primary shadow-sm'
          : hasIssues
            ? 'bg-amber-950/40 text-amber-300 border-amber-800/80 hover:bg-amber-950/70'
            : 'bg-secondary text-muted-foreground hover:text-foreground border-border'
      }`;
    });
  }

  const s = slides[idx];
  if (!s) return;

  // Update Info labels
  const tplInfo = document.getElementById('ht-verify-tpl-info');
  const genInfo = document.getElementById('ht-verify-gen-info');
  if (tplInfo) tplInfo.innerText = `${s.source_template || 'Template'} (Slide #${(s.source_slide_index ?? 0) + 1})`;
  if (genInfo) genInfo.innerText = `Slide ${s.slide_number}: ${s.title || ''}`;

  // Update Template Image
  const tplImg = document.getElementById('ht-verify-tpl-img');
  const tplPh = document.getElementById('ht-verify-tpl-placeholder');
  if (tplImg && tplPh) {
    if (s.template_screenshot) {
      tplImg.src = s.template_screenshot;
      tplImg.classList.remove('hidden');
      tplPh.classList.add('hidden');
    } else {
      tplImg.classList.add('hidden');
      tplPh.classList.remove('hidden');
    }
  }

  // Update Generated Image
  const genImg = document.getElementById('ht-verify-gen-img');
  const genPh = document.getElementById('ht-verify-gen-placeholder');
  if (genImg && genPh) {
    if (s.generated_screenshot) {
      genImg.src = s.generated_screenshot;
      genImg.classList.remove('hidden');
      genPh.classList.add('hidden');
    } else {
      genImg.classList.add('hidden');
      genPh.classList.remove('hidden');
    }
  }

  // Update Detected Issues List
  const issuesList = document.getElementById('ht-verify-issues-list');
  const issueCount = document.getElementById('ht-verify-issue-count');
  const issues = s.detected_issues || [];
  if (issueCount) issueCount.innerText = `${issues.length} issue(s)`;
  if (issuesList) {
    if (issues.length > 0) {
      issuesList.innerHTML = issues.map(iss => `
        <div class="p-2 rounded bg-amber-950/30 border border-amber-800/50 text-amber-200 flex items-start gap-2">
          <i data-lucide="alert-circle" class="w-3.5 h-3.5 text-amber-400 shrink-0 mt-0.5"></i>
          <span class="leading-relaxed">${escapeHtml(iss)}</span>
        </div>
      `).join('');
    } else {
      issuesList.innerHTML = `
        <div class="p-2.5 rounded bg-emerald-950/20 border border-emerald-800/40 text-emerald-300 flex items-center gap-2">
          <i data-lucide="check-circle-2" class="w-3.5 h-3.5 text-emerald-400 shrink-0"></i>
          <span>No missing elements, placeholder text, or layout issues detected on this slide.</span>
        </div>
      `;
    }
  }

  // Update Actions List
  const actionsList = document.getElementById('ht-verify-actions-list');
  const actionCount = document.getElementById('ht-verify-action-count');
  const actions = s.edit_structure?.actions || [];
  if (actionCount) actionCount.innerText = `${actions.length} action(s)`;
  if (actionsList) {
    if (actions.length > 0) {
      actionsList.innerHTML = actions.map((act, aIdx) => {
        const isDelete = ['delete_shape', 'remove_shape', 'delete_shapes', 'remove_shapes', 'delete_element', 'remove_element', 'remove_formula', 'delete_formula'].includes(act.action);
        const iconName = isDelete ? 'trash-2' : 'wrench';
        const badgeColor = isDelete ? 'bg-rose-950/40 border-rose-800/60 text-rose-200' : 'bg-sky-950/30 border-sky-800/50 text-sky-200';
        const iconColor = isDelete ? 'text-rose-400' : 'text-sky-400';

        let desc = act.action;
        if (['delete_shape', 'remove_shape', 'delete_element', 'remove_element'].includes(act.action)) {
          desc = `Delete shape #${act.shape_index}${act.formula_text ? ` (Formula: "${escapeHtml(act.formula_text)}")` : ''}`;
        } else if (['remove_formula', 'delete_formula'].includes(act.action)) {
          desc = `Delete unreplaced formula${act.formula_text ? `: "${escapeHtml(act.formula_text)}"` : ' equation from template'}`;
        } else if (act.action === 'remove_shapes' || act.action === 'delete_shapes') {
          desc = `Delete shape(s) #${(act.shape_indices || []).join(', #')}`;
        } else if (act.action === 'update_text') {
          desc = `Update shape #${act.shape_index} text: "${escapeHtml(act.new_text || '')}"`;
        } else if (act.action === 'update_table') {
          desc = `Update table shape #${act.shape_index} cell contents`;
        } else if (act.action === 'update_notes') {
          desc = `Update speaker notes`;
        }
        return `
          <div class="p-2 rounded ${badgeColor} border flex items-start justify-between gap-2">
            <div class="flex items-start gap-2">
              <i data-lucide="${iconName}" class="w-3.5 h-3.5 ${iconColor} shrink-0 mt-0.5"></i>
              <span class="leading-relaxed">${desc}</span>
            </div>
            <button type="button" onclick="removeVerificationAction(${idx}, ${aIdx})" class="text-muted-foreground hover:text-rose-400 text-[10px] p-0.5" title="Dismiss this action">
              ✕
            </button>
          </div>
        `;
      }).join('');
    } else {
      actionsList.innerHTML = `
        <div class="p-2.5 rounded bg-secondary/30 border border-border/40 text-muted-foreground flex items-center gap-2">
          <i data-lucide="check" class="w-3.5 h-3.5 text-muted-foreground/60 shrink-0"></i>
          <span>No healing modifications required for this slide.</span>
        </div>
      `;
    }
  }

  // Render Conversation Stream
  renderVerificationChatStream();

  if (window.lucide) lucide.createIcons();
}

function removeVerificationAction(slideIdx, actionIdx) {
  const slides = currentHumanTouchData.slides || [];
  if (slides[slideIdx] && slides[slideIdx].edit_structure && slides[slideIdx].edit_structure.actions) {
    slides[slideIdx].edit_structure.actions.splice(actionIdx, 1);
    // Recompute aggregated_actions
    currentHumanTouchData.aggregated_actions = slides.flatMap(s => s.edit_structure?.actions || []);
    currentHumanTouchData.total_actions_planned = currentHumanTouchData.aggregated_actions.length;
    selectVerifySlide(slideIdx);
  }
}

function renderVerificationChatStream() {
  const stream = document.getElementById('ht-verify-chat-stream');
  if (!stream) return;
  const history = currentHumanTouchData.conversation_history || [];

  let html = `
    <div class="p-2 rounded bg-secondary/40 text-muted-foreground text-[11px]">
      👋 <strong class="text-foreground">Verification Agent:</strong> I inspected the chosen template slide against the generated slide. Tell me if any elements, badges, or texts are missing or need adjusting.
    </div>
  `;

  history.forEach(item => {
    if (item.role === 'user') {
      html += `
        <div class="p-2 rounded bg-primary/10 border border-primary/20 text-foreground text-[11px] ml-4">
          <strong class="text-primary">You:</strong> ${escapeHtml(item.content)}
        </div>
      `;
    } else if (item.role === 'assistant') {
      html += `
        <div class="p-2 rounded bg-secondary/50 border border-border/60 text-foreground text-[11px] mr-4">
          <strong class="text-emerald-400">Verification Agent:</strong> ${escapeHtml(item.content)}
        </div>
      `;
    }
  });

  stream.innerHTML = html;
  stream.scrollTop = stream.scrollHeight;
}

function skipVerificationEdits() {
  if (currentHumanTouchData) {
    currentHumanTouchData.aggregated_actions = [];
    if (Array.isArray(currentHumanTouchData.slides)) {
      currentHumanTouchData.slides.forEach(s => {
        if (s.edit_structure) {
          s.edit_structure.actions = [];
          s.edit_structure.summary_of_changes = 'Verification edits skipped by user';
        }
      });
    }
    currentHumanTouchData.total_actions_planned = 0;
  }
  continueHumanTouchJob();
}

function formatHumanTouchJson() {
  const jsonArea = document.getElementById('ht-raw-json-textarea');
  if (!jsonArea) return;
  try {
    const parsed = JSON.parse(jsonArea.value);
    jsonArea.value = JSON.stringify(parsed, null, 2);
    showToast('JSON formatted successfully', 'success');
  } catch (err) {
    showToast(`Invalid JSON: ${err.message}`, 'error');
  }
}

function renderHumanTouchVisual(step, data) {
  const container = document.getElementById('ht-cards-list');
  if (!container) return;
  container.innerHTML = '';

  const docTitleInput = document.getElementById('ht-doc-title-input');
  if (docTitleInput && data.document_title) {
    docTitleInput.value = data.document_title;
  }

  if (step === 'extract') {
    const sections = Array.isArray(data.sections) ? data.sections : [];
    if (sections.length === 0) {
      container.innerHTML = '<div class="text-center py-6 text-muted-foreground italic">No slide sections found. Click "+ Add New Slide / Section" to add one.</div>';
      return;
    }

    sections.forEach((sec, idx) => {
      const card = document.createElement('div');
      card.className = 'ht-card p-3 rounded-lg bg-secondary/30 border border-border/70 space-y-2.5 transition';
      card.dataset.index = idx;

      const paragraphsText = Array.isArray(sec.paragraphs) ? sec.paragraphs.join('\n') : (sec.paragraphs || '');
      const bulletsText = Array.isArray(sec.bullets) ? sec.bullets.map(b => b.startsWith('•') || b.startsWith('-') ? b : `• ${b}`).join('\n') : (sec.bullets || '');

      card.innerHTML = `
        <div class="flex items-center justify-between pb-1.5 border-b border-border/50">
          <div class="flex items-center gap-2">
            <span class="w-5 h-5 rounded-full bg-primary/20 text-primary flex items-center justify-center font-mono text-[10px] font-bold">
              ${idx + 1}
            </span>
            <span class="font-medium text-foreground text-xs">Slide Section ${idx + 1}</span>
          </div>
          <button type="button" onclick="deleteHumanTouchCard(${idx})"
            class="text-muted-foreground hover:text-rose-400 p-1 rounded transition" title="Delete section">
            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
          </button>
        </div>

        <div class="space-y-1">
          <label class="font-medium text-[11px] text-muted-foreground">Section Title:</label>
          <input type="text" class="ht-sec-title shadcn-input text-xs font-semibold" value="${(sec.title || '').replace(/"/g, '&quot;')}" placeholder="Slide Title">
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-2.5 pt-1">
          <div class="space-y-1">
            <label class="font-medium text-[11px] text-muted-foreground">Paragraphs (one per line):</label>
            <textarea class="ht-sec-paragraphs shadcn-input font-mono text-xs leading-relaxed resize-y" rows="3" placeholder="Detailed paragraphs or explanations...">${paragraphsText}</textarea>
          </div>
          <div class="space-y-1">
            <label class="font-medium text-[11px] text-muted-foreground">Bullet Points (one per line):</label>
            <textarea class="ht-sec-bullets shadcn-input font-mono text-xs leading-relaxed resize-y" rows="3" placeholder="• Key bullet point 1&#10;• Key bullet point 2">${bulletsText}</textarea>
          </div>
        </div>
      `;
      container.appendChild(card);
    });
  } else if (step === 'restructure') {
    const slides = Array.isArray(data.slides) ? data.slides : [];
    if (slides.length === 0) {
      container.innerHTML = '<div class="text-center py-6 text-muted-foreground italic">No storyboard slides found. Click "+ Add New Slide / Section" to add one.</div>';
      return;
    }

    slides.forEach((sld, idx) => {
      const card = document.createElement('div');
      card.className = 'ht-card p-3 rounded-lg bg-secondary/30 border border-border/70 space-y-2.5 transition';
      card.dataset.index = idx;

      const quad = sld.quadrant || 'TR';
      const sType = sld.slide_type || 'Theory / Definition';

      let bulletsText = '';
      if (Array.isArray(sld.rewritten_bullets)) {
        bulletsText = sld.rewritten_bullets.join('\n');
      } else if (Array.isArray(sld.animation_steps)) {
        bulletsText = sld.animation_steps.map(st => `${st.indicator || '①'} ${st.text || ''}`).join('\n');
      }

      const formulas = (sld.mathematical_elements && sld.mathematical_elements.formulas) || '';
      const callout = (sld.visual_annotations && sld.visual_annotations.teacher_callouts) || '';
      const notes = sld.speaker_notes || '';

      card.innerHTML = `
        <div class="flex items-center justify-between pb-1.5 border-b border-border/50">
          <div class="flex items-center gap-2">
            <span class="w-5 h-5 rounded-full bg-amber-500/20 text-amber-400 flex items-center justify-center font-mono text-[10px] font-bold">
              ${idx + 1}
            </span>
            <span class="font-medium text-foreground text-xs">Slide ${idx + 1} (${quad})</span>
            <span class="text-[10px] font-mono px-1.5 py-0.5 rounded bg-secondary border border-border text-muted-foreground">${sType}</span>
          </div>
          <button type="button" onclick="deleteHumanTouchCard(${idx})"
            class="text-muted-foreground hover:text-rose-400 p-1 rounded transition" title="Delete slide">
            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
          </button>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-2">
          <div class="md:col-span-2 space-y-1">
            <label class="font-medium text-[11px] text-muted-foreground">Slide Title:</label>
            <input type="text" class="ht-sld-title shadcn-input text-xs font-semibold" value="${(sld.title || '').replace(/"/g, '&quot;')}" placeholder="Slide Title">
          </div>
          <div class="space-y-1">
            <label class="font-medium text-[11px] text-muted-foreground">Quadrant / Position:</label>
            <select class="ht-sld-quadrant shadcn-input text-xs cursor-pointer font-mono">
              <option value="TR" ${quad === 'TR' ? 'selected' : ''}>TR (Top-Right)</option>
              <option value="TL" ${quad === 'TL' ? 'selected' : ''}>TL (Top-Left)</option>
              <option value="BR" ${quad === 'BR' ? 'selected' : ''}>BR (Bottom-Right)</option>
              <option value="BL" ${quad === 'BL' ? 'selected' : ''}>BL (Bottom-Left)</option>
            </select>
          </div>
        </div>

        <div class="space-y-1">
          <label class="font-medium text-[11px] text-muted-foreground">Core Concept (1-line punchy summary):</label>
          <input type="text" class="ht-sld-concept shadcn-input text-xs" value="${(sld.core_concept || '').replace(/"/g, '&quot;')}" placeholder="One-line conceptual essence...">
        </div>

        <div class="space-y-1">
          <label class="font-medium text-[11px] text-muted-foreground">Numbered Steps & Rewritten Bullets (①, ②, ③...):</label>
          <textarea class="ht-sld-bullets shadcn-input font-mono text-xs leading-relaxed resize-y" rows="3" placeholder="① First step...&#10;② Second step...">${bulletsText}</textarea>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
          <div class="space-y-1">
            <label class="font-medium text-[11px] text-muted-foreground">Formulas / LaTeX:</label>
            <input type="text" class="ht-sld-formulas shadcn-input font-mono text-xs" value="${formulas.replace(/"/g, '&quot;')}" placeholder="$E=mc^2$">
          </div>
          <div class="space-y-1">
            <label class="font-medium text-[11px] text-muted-foreground">Teacher Callout / Invariant:</label>
            <input type="text" class="ht-sld-callout shadcn-input text-xs" value="${callout.replace(/"/g, '&quot;')}" placeholder="نکته کلیدی: ...">
          </div>
        </div>

        <div class="space-y-1">
          <label class="font-medium text-[11px] text-muted-foreground">Speaker Notes:</label>
          <textarea class="ht-sld-notes shadcn-input text-xs resize-y" rows="2" placeholder="Explanatory notes for presenter...">${notes}</textarea>
        </div>
      `;
      container.appendChild(card);
    });
  }

  if (window.lucide) lucide.createIcons();
}

function addHumanTouchCard() {
  currentHumanTouchData = collectHumanTouchDataFromVisual();
  if (currentHumanTouchStep === 'extract') {
    if (!Array.isArray(currentHumanTouchData.sections)) {
      currentHumanTouchData.sections = [];
    }
    currentHumanTouchData.sections.push({
      title: `New Section ${currentHumanTouchData.sections.length + 1}`,
      paragraphs: ['Additional context or details.'],
      bullets: ['Key takeaway point']
    });
  } else if (currentHumanTouchStep === 'restructure') {
    if (!Array.isArray(currentHumanTouchData.slides)) {
      currentHumanTouchData.slides = [];
    }
    const idx = currentHumanTouchData.slides.length + 1;
    const quads = ['TR', 'TL', 'BR', 'BL'];
    currentHumanTouchData.slides.push({
      slide_number: idx,
      title: `New Slide ${idx}`,
      quadrant: quads[(idx - 1) % 4],
      slide_type: 'Theory / Definition',
      core_concept: 'Key concept definition',
      animation_steps: [{ step: 1, indicator: '①', text: 'First main step' }],
      rewritten_bullets: ['① First main step with key insight'],
      mathematical_elements: { definitions: '', formulas: '', sets: '' },
      visual_annotations: { teacher_callouts: 'نکته کلیدی' },
      speaker_notes: ''
    });
  }
  renderHumanTouchVisual(currentHumanTouchStep, currentHumanTouchData);
}

function deleteHumanTouchCard(idx) {
  currentHumanTouchData = collectHumanTouchDataFromVisual();
  if (currentHumanTouchStep === 'extract' && Array.isArray(currentHumanTouchData.sections)) {
    currentHumanTouchData.sections.splice(idx, 1);
  } else if (currentHumanTouchStep === 'restructure' && Array.isArray(currentHumanTouchData.slides)) {
    currentHumanTouchData.slides.splice(idx, 1);
  }
  renderHumanTouchVisual(currentHumanTouchStep, currentHumanTouchData);
}

function collectHumanTouchDataFromVisual() {
  const docTitleInput = document.getElementById('ht-doc-title-input');
  const docTitle = docTitleInput ? docTitleInput.value.trim() : (currentHumanTouchData.document_title || 'Presentation');
  const cards = document.querySelectorAll('#ht-cards-list .ht-card');

  if (currentHumanTouchStep === 'extract') {
    const sections = [];
    cards.forEach((c) => {
      const title = c.querySelector('.ht-sec-title')?.value.trim() || 'Section';
      const rawParas = c.querySelector('.ht-sec-paragraphs')?.value || '';
      const rawBullets = c.querySelector('.ht-sec-bullets')?.value || '';

      const paragraphs = rawParas.split('\n').map(p => p.trim()).filter(Boolean);
      const bullets = rawBullets.split('\n').map(b => b.trim()).filter(Boolean).map(b => b.replace(/^[•\-*]\s*/, ''));

      sections.push({
        title: title,
        paragraphs: paragraphs,
        bullets: bullets
      });
    });

    return {
      document_title: docTitle,
      total_sections: sections.length,
      sections: sections,
      sources_summary: currentHumanTouchData.sources_summary || [],
      raw_text: currentHumanTouchData.raw_text || ''
    };
  } else if (currentHumanTouchStep === 'restructure') {
    const slides = [];
    const stepIcons = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧'];

    cards.forEach((c, idx) => {
      const title = c.querySelector('.ht-sld-title')?.value.trim() || `Slide ${idx + 1}`;
      const quad = c.querySelector('.ht-sld-quadrant')?.value || 'TR';
      const concept = c.querySelector('.ht-sld-concept')?.value.trim() || '';
      const rawBullets = c.querySelector('.ht-sld-bullets')?.value || '';
      const formulas = c.querySelector('.ht-sld-formulas')?.value.trim() || '';
      const callout = c.querySelector('.ht-sld-callout')?.value.trim() || '';
      const notes = c.querySelector('.ht-sld-notes')?.value.trim() || '';

      const lines = rawBullets.split('\n').map(l => l.trim()).filter(Boolean);
      const animSteps = lines.map((line, lIdx) => {
        const icon = stepIcons[lIdx % stepIcons.length];
        const clean = line.replace(/^[①②③④⑤⑥⑦⑧•\-*]\s*/, '');
        return { step: lIdx + 1, indicator: icon, text: clean };
      });

      slides.push({
        slide_number: idx + 1,
        title: title,
        quadrant: quad,
        slide_type: 'Theory / Definition',
        core_concept: concept,
        animation_steps: animSteps,
        mathematical_elements: { definitions: '', formulas: formulas, sets: '' },
        visual_annotations: { teacher_callouts: callout, key_invariants: '' },
        rewritten_bullets: lines,
        speaker_notes: notes
      });
    });

    return {
      document_title: docTitle,
      total_slides: slides.length,
      slides: slides
    };
  } else if (currentHumanTouchStep === 'verify') {
    return {
      ...currentHumanTouchData,
      active_slide_index: currentVerifySlideIdx
    };
  }

  return currentHumanTouchData;
}

async function rerunHumanTouchProcess() {
  const promptInput = document.getElementById('ht-reply-prompt-input');
  const prompt = promptInput ? promptInput.value.trim() : '';

  if (!prompt) {
    showToast('Please type an instruction or pick a suggestion chip to rerun.', 'warning');
    if (promptInput) promptInput.focus();
    return;
  }

  const btnRerun = document.getElementById('btn-ht-rerun');
  const statusSpan = document.getElementById('ht-rerun-status');

  if (currentHumanTouchView === 'visual') {
    currentHumanTouchData = collectHumanTouchDataFromVisual();
  } else {
    try {
      const jsonArea = document.getElementById('ht-raw-json-textarea');
      if (jsonArea && jsonArea.value.trim()) {
        currentHumanTouchData = JSON.parse(jsonArea.value.trim());
      }
    } catch (err) {
      showToast('Invalid JSON syntax; please fix before rerunning.', 'error');
      return;
    }
  }

  btnRerun.disabled = true;
  if (currentHumanTouchStep === 'verify') {
    statusSpan.innerText = 'Verification Agent analyzing conversation and refining slide actions...';
    appendGenLog(`[Verification Chat] User instruction: "${prompt}"`);
  } else {
    statusSpan.innerText = 'AI is rerunning step with your feedback...';
    appendGenLog(`[Human Touch] Requested AI rerun for ${currentHumanTouchStep}: "${prompt}"`);
  }

  try {
    const res = await fetch('/api/generator/human-action', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        job_id: currentHumanTouchJobId,
        action: 'rerun',
        step: currentHumanTouchStep,
        prompt: prompt,
        data: currentHumanTouchData
      })
    });

    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || 'Failed to signal rerun.');
    }
    if (promptInput) {
      promptInput.value = '';
    }
  } catch (err) {
    statusSpan.innerText = `Error: ${err.message}`;
    btnRerun.disabled = false;
    showToast(`Rerun error: ${err.message}`, 'error');
  }
}

async function continueHumanTouchJob() {
  if (currentHumanTouchView === 'visual') {
    currentHumanTouchData = collectHumanTouchDataFromVisual();
  } else {
    try {
      const jsonArea = document.getElementById('ht-raw-json-textarea');
      if (jsonArea && jsonArea.value.trim()) {
        currentHumanTouchData = JSON.parse(jsonArea.value.trim());
      }
    } catch (err) {
      showToast('Invalid JSON syntax; please correct before continuing.', 'error');
      return;
    }
  }

  const btnCont = document.getElementById('btn-ht-continue');
  btnCont.disabled = true;

  try {
    const res = await fetch('/api/generator/human-action', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        job_id: currentHumanTouchJobId,
        action: 'continue',
        step: currentHumanTouchStep,
        data: currentHumanTouchData
      })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    closeHumanTouchModal();
    appendGenLog(`[Human Touch] Step '${currentHumanTouchStep}' approved by user. Resuming pipeline...`);
    showToast('Changes approved. Continuing generation pipeline...', 'success');
  } catch (err) {
    showToast(`Could not resume pipeline: ${err.message}`, 'error');
  } finally {
    btnCont.disabled = false;
  }
}

function cancelHumanTouchJob() {
  if (!confirm('Are you sure you want to cancel the generation pipeline?')) return;
  if (currentHumanTouchJobId) {
    fetch('/api/generator/human-action', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        job_id: currentHumanTouchJobId,
        action: 'cancel'
      })
    });
  }
  closeHumanTouchModal();
  closeFontFallbackModal();
  setSystemStatus('CANCELLED');
  appendGenLog('[Human Touch] Pipeline cancelled by user.');
}

function closeHumanTouchModal() {
  const modal = document.getElementById('human-touch-modal');
  if (modal) modal.classList.add('hidden');
}

function closeFontFallbackModal() {
  const modal = document.getElementById('font-fallback-modal');
  if (modal) modal.classList.add('hidden');
}

// -------------------------------------------------------------
// 1.5. TEMPLATE FONT VERIFICATION & SEARCHABLE FALLBACK RESOLVER
// -------------------------------------------------------------
let currentFontFallbackPayload = null;
let currentConfiguredFontFallbacks = {};
let fontFallbackItems = []; // [{ missingFont, rec, selectedValue, customValue, query, category }]
let fontFallbackSystemCatalog = {
  families: [],
  persian_fonts: [],
  latin_fonts: []
};

async function checkGeneratorTemplateFonts() {
  const select = document.getElementById('gen-template-select');
  const statusDiv = document.getElementById('gen-template-font-status');
  if (!select || !statusDiv) return;

  const tpl = select.value.trim();

  try {
    const url = tpl ? `/api/fonts/verify?template_name=${encodeURIComponent(tpl)}` : `/api/fonts/verify`;
    const res = await fetch(url);
    const data = await res.json();

    if (!data.success) {
      statusDiv.classList.add('hidden');
      return;
    }

    statusDiv.classList.remove('hidden');

    if (data.all_installed) {
      statusDiv.className = 'mt-1.5 text-[11px] p-2 rounded border border-emerald-500/30 bg-emerald-500/10 text-emerald-400 flex items-center justify-between';
      statusDiv.innerHTML = `
        <div class="flex items-center gap-1.5">
          <i data-lucide="check-circle" class="w-3.5 h-3.5 text-emerald-400 shrink-0"></i>
          <span>All template fonts verified on system: <strong class="font-mono text-foreground">${data.installed_fonts.join(', ') || 'Standard'}</strong></span>
        </div>
      `;
    } else {
      statusDiv.className = 'mt-1.5 text-[11px] p-2 rounded border border-amber-500/30 bg-amber-500/10 text-amber-400 flex items-center justify-between flex-wrap gap-1.5';
      const missingList = data.missing_fonts.join(', ');
      statusDiv.innerHTML = `
        <div class="flex items-center gap-1.5">
          <i data-lucide="alert-triangle" class="w-3.5 h-3.5 text-amber-400 shrink-0"></i>
          <span>Missing fonts on system: <strong class="font-mono text-foreground">${missingList}</strong></span>
        </div>
        <div class="flex items-center gap-1.5">
          <span class="text-[10px] text-muted-foreground font-mono bg-background/50 px-1.5 py-0.5 rounded border border-border/50">
            Step 1.2
          </span>
          <button type="button" onclick="openPreConfigFontFallbackModal()"
            class="shadcn-btn shadcn-btn-outline text-[11px] h-6 px-2 py-0 gap-1 text-amber-300 border-amber-500/40 hover:bg-amber-500/20 cursor-pointer">
            <i data-lucide="search" class="w-3 h-3"></i> Search &amp; Pick Fallback
          </button>
        </div>
      `;
    }
    if (window.lucide) lucide.createIcons();
  } catch (err) {
    statusDiv.classList.add('hidden');
  }
}

function openPreConfigFontFallbackModal() {
  const select = document.getElementById('gen-template-select');
  const tpl = select ? select.value.trim() : '';
  const url = tpl ? `/api/fonts/verify?template_name=${encodeURIComponent(tpl)}` : `/api/fonts/verify`;
  fetch(url)
    .then(r => r.json())
    .then(data => {
      if (data.success && data.missing_fonts && data.missing_fonts.length > 0) {
        openFontFallbackModal({
          job_id: null,
          step: 'font_fallback',
          data: data
        });
      } else {
        showToast('All template fonts are already installed on system!', 'success');
      }
    })
    .catch(err => {
      showToast(`Could not load fonts: ${err.message}`, 'error');
    });
}

async function rescanSystemFontsAction() {
  const btn = document.getElementById('btn-ff-rescan');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<i data-lucide="loader-2" class="w-3 h-3 animate-spin"></i> Rescanning...`;
    if (window.lucide) lucide.createIcons();
  }

  try {
    const res = await fetch('/api/fonts/system/rescan', { method: 'POST' });
    const d = await res.json();
    if (d.success) {
      fontFallbackSystemCatalog = {
        families: d.families || [],
        persian_fonts: d.persian_fonts || [],
        latin_fonts: d.latin_fonts || []
      };

      // Re-verify current template fonts
      const select = document.getElementById('gen-template-select');
      const tpl = select ? select.value.trim() : '';
      const vUrl = tpl ? `/api/fonts/verify?template_name=${encodeURIComponent(tpl)}&rescan=true` : `/api/fonts/verify?rescan=true`;
      const vRes = await fetch(vUrl);
      const vData = await vRes.json();

      if (vData.success) {
        currentFontFallbackPayload = vData;
        const missingFonts = vData.missing_fonts || [];
        const recommendations = vData.recommendations || {};
        const currentFallbacks = vData.current_fallbacks || {};

        fontFallbackItems = missingFonts.map(mf => {
          const rec = recommendations[mf] || currentFallbacks[mf] || 'Segoe UI';
          const existingConfig = currentConfiguredFontFallbacks[mf] || currentFallbacks[mf];
          return {
            missingFont: mf,
            rec: rec,
            selectedValue: existingConfig || rec,
            customValue: '',
            query: '',
            category: 'all'
          };
        });

        renderFontFallbackList();
        checkGeneratorTemplateFonts();
        showToast(`Fonts rescanned: ${d.families.length} fonts discovered (${d.persian_fonts.length} Persian, ${d.latin_fonts.length} Latin)`, 'success');
      }
    } else {
      showToast('Failed to rescan fonts: ' + (d.error || 'Unknown error'), 'error');
    }
  } catch (err) {
    showToast('Failed to rescan fonts: ' + err.message, 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `<i data-lucide="refresh-cw" class="w-3 h-3"></i> Rescan Fonts`;
      if (window.lucide) lucide.createIcons();
    }
  }
}

function openFontFallbackModal(payload) {
  currentHumanTouchJobId = payload.job_id || null;
  currentHumanTouchStep = payload.step || 'font_fallback';
  currentFontFallbackPayload = payload.data || {};

  const missingFonts = currentFontFallbackPayload.missing_fonts || [];
  const recommendations = currentFontFallbackPayload.recommendations || {};
  const currentFallbacks = currentFontFallbackPayload.current_fallbacks || {};

  fontFallbackSystemCatalog = {
    families: currentFontFallbackPayload.system_fonts || [],
    persian_fonts: currentFontFallbackPayload.available_persian || [],
    latin_fonts: currentFontFallbackPayload.available_latin || []
  };

  // If system fonts empty, fetch from API
  if (!fontFallbackSystemCatalog.families || fontFallbackSystemCatalog.families.length === 0) {
    fetch('/api/fonts/system')
      .then(r => r.json())
      .then(d => {
        if (d.success) {
          fontFallbackSystemCatalog = {
            families: d.families || [],
            persian_fonts: d.persian_fonts || [],
            latin_fonts: d.latin_fonts || []
          };
          renderFontFallbackList();
        }
      })
      .catch(() => {});
  }

  fontFallbackItems = missingFonts.map(mf => {
    const rec = recommendations[mf] || currentFallbacks[mf] || 'Segoe UI';
    const existingConfig = currentConfiguredFontFallbacks[mf] || currentFallbacks[mf];
    return {
      missingFont: mf,
      rec: rec,
      selectedValue: existingConfig || rec,
      customValue: '',
      query: '',
      category: 'all'
    };
  });

  renderFontFallbackList();

  const modal = document.getElementById('font-fallback-modal');
  if (modal) modal.classList.remove('hidden');
  if (window.lucide) lucide.createIcons();

  showToast(`Step 1.2: ${missingFonts.length} template font(s) need fallback selection`, 'warning');
}

function renderFontFallbackList() {
  const container = document.getElementById('ff-missing-fonts-list');
  if (!container) return;
  container.innerHTML = '';

  if (fontFallbackItems.length === 0) {
    container.innerHTML = `
      <div class="p-4 rounded-lg bg-secondary/30 border border-border/70 text-center text-muted-foreground text-xs">
        No missing fonts detected. All template typography is verified on this system.
      </div>
    `;
    return;
  }

  // Batch action bar if multiple fonts missing
  if (fontFallbackItems.length > 1) {
    const batchDiv = document.createElement('div');
    batchDiv.className = 'p-2.5 rounded-lg border border-primary/30 bg-primary/5 flex items-center justify-between flex-wrap gap-2 text-xs mb-2';

    const quickChoices = [
      'IRANYekanXFaNum',
      'IRANYekanXFaNum Heavy',
      'Vazirmatn',
      'B Nazanin',
      'Segoe UI',
      'Calibri',
      'Arial',
      'Tahoma'
    ].filter(f => fontFallbackSystemCatalog.families.some(sf => sf.toLowerCase() === f.toLowerCase()));

    const batchList = quickChoices.length > 0 ? quickChoices : fontFallbackSystemCatalog.families.slice(0, 10);
    const batchOptions = batchList.map(f => `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`).join('');

    batchDiv.innerHTML = `
      <div class="flex items-center gap-1.5 font-medium text-foreground">
        <i data-lucide="wand-2" class="w-3.5 h-3.5 text-primary"></i>
        <span>Quick Apply to All Missing Fonts:</span>
      </div>
      <div class="flex items-center gap-1.5 flex-1 min-w-[220px]">
        <select id="ff-batch-select" class="shadcn-input text-xs flex-1 cursor-pointer">
          ${batchOptions}
        </select>
        <button type="button" onclick="applyBatchFontFallback()" class="shadcn-btn shadcn-btn-secondary text-xs px-2.5 py-1 shrink-0 gap-1 font-medium cursor-pointer">
          Apply All
        </button>
      </div>
    `;
    container.appendChild(batchDiv);
  }

  // Render each item
  fontFallbackItems.forEach((item, idx) => {
    const card = document.createElement('div');
    card.id = `ff-card-${idx}`;
    card.className = 'p-3.5 rounded-lg border border-border/70 bg-secondary/30 space-y-2.5 transition-all';
    card.innerHTML = getFontFallbackItemMarkup(item, idx);
    container.appendChild(card);
  });

  if (window.lucide) lucide.createIcons();
}

function getFontFallbackItemMarkup(item, idx) {
  const { optionsHtml, statsText } = buildFontFallbackOptions(item);
  const curDisplay = item.selectedValue === '__custom__' ? (item.customValue || 'Custom') : item.selectedValue;

  return `
    <!-- Top Header: Missing font name + active chosen font -->
    <div class="flex items-center justify-between flex-wrap gap-1.5">
      <div class="flex items-center gap-2">
        <span class="font-bold text-foreground text-xs font-mono bg-background px-2 py-0.5 rounded border border-border/60">${escapeHtml(item.missingFont)}</span>
        <span class="px-2 py-0.5 rounded bg-rose-500/10 text-rose-400 border border-rose-500/30 text-[10px] font-mono">Missing on System</span>
      </div>
      <div class="text-[11px] text-muted-foreground font-mono">
        Active Fallback: <span class="font-semibold text-primary" id="ff-cur-display-${idx}">${escapeHtml(curDisplay)}</span>
      </div>
    </div>

    <!-- Search Input & Script Filter Tabs -->
    <div class="flex items-center gap-1.5 flex-wrap sm:flex-nowrap">
      <div class="relative flex-1 min-w-[200px]">
        <i data-lucide="search" class="w-3.5 h-3.5 text-muted-foreground absolute left-2.5 top-2.5 pointer-events-none"></i>
        <input type="text" id="ff-search-${idx}" value="${escapeHtml(item.query)}"
          placeholder="🔍 Search among 1,000+ fonts (e.g. Yekan, Segoe, Nazanin, Calibri)..."
          class="shadcn-input text-xs pl-8 pr-7 w-full font-mono"
          oninput="onFontSearchInput(${idx}, this.value)"
          onkeydown="if(event.key==='Enter'){event.preventDefault(); onFontSearchEnter(${idx});}">
        ${item.query ? `
          <button type="button" onclick="clearFontSearch(${idx})"
            class="absolute right-2 top-2 text-muted-foreground hover:text-foreground text-xs h-4 w-4 flex items-center justify-center cursor-pointer" title="Clear search">
            ✕
          </button>
        ` : ''}
      </div>

      <!-- Script Filter Pills -->
      <div class="flex items-center p-0.5 rounded-md bg-secondary border border-border text-[10px] shrink-0">
        <button type="button" onclick="setFontCategory(${idx}, 'all')"
          class="px-2 py-0.5 rounded font-medium transition cursor-pointer ${item.category === 'all' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-muted-foreground hover:text-foreground'}">
          All (${fontFallbackSystemCatalog.families.length || 0})
        </button>
        <button type="button" onclick="setFontCategory(${idx}, 'persian')"
          class="px-2 py-0.5 rounded font-medium transition cursor-pointer ${item.category === 'persian' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-muted-foreground hover:text-foreground'}">
          Persian (${fontFallbackSystemCatalog.persian_fonts.length || 0})
        </button>
        <button type="button" onclick="setFontCategory(${idx}, 'latin')"
          class="px-2 py-0.5 rounded font-medium transition cursor-pointer ${item.category === 'latin' ? 'bg-primary text-primary-foreground shadow-sm' : 'text-muted-foreground hover:text-foreground'}">
          Latin (${fontFallbackSystemCatalog.latin_fonts.length || 0})
        </button>
      </div>
    </div>

    <!-- Select Dropdown & Custom Font Input -->
    <div class="space-y-1.5">
      <select id="ff-select-${idx}" data-source-font="${escapeHtml(item.missingFont)}"
        onchange="onFontFallbackSelectChange(${idx}, this.value)"
        class="shadcn-input w-full text-xs cursor-pointer font-medium">
        ${optionsHtml}
      </select>

      <input type="text" id="ff-custom-${idx}" value="${escapeHtml(item.customValue || '')}"
        placeholder="Type custom font family name (e.g. 'Vazirmatn', 'IRANYekanXFaNum', 'Segoe UI')..."
        oninput="onFontCustomInput(${idx}, this.value)"
        class="shadcn-input text-xs ${item.selectedValue === '__custom__' ? '' : 'hidden'} w-full font-mono">
    </div>

    <!-- Search Stats & Enter Hint -->
    <div class="flex items-center justify-between text-[10px] text-muted-foreground font-mono">
      <span id="ff-stats-${idx}">${statsText}</span>
      <span class="opacity-80">Tip: Type query &amp; press Enter to select</span>
    </div>
  `;
}

function buildFontFallbackOptions(item) {
  const query = (item.query || '').trim().toLowerCase();
  const cat = item.category || 'all';

  let candidates = [];
  if (cat === 'persian') {
    candidates = fontFallbackSystemCatalog.persian_fonts || [];
  } else if (cat === 'latin') {
    candidates = fontFallbackSystemCatalog.latin_fonts || [];
  } else {
    candidates = fontFallbackSystemCatalog.families || [];
  }

  if (query) {
    candidates = candidates.filter(f => f.toLowerCase().includes(query));
  }

  let optionsHtml = '';

  // 1. Recommended Font
  if (item.rec && (!query || item.rec.toLowerCase().includes(query))) {
    optionsHtml += `<option value="${escapeHtml(item.rec)}" ${item.selectedValue === item.rec ? 'selected' : ''}>✨ Recommended: ${escapeHtml(item.rec)}</option>`;
  }

  // 2. Custom typed query option if query present
  if (query && !candidates.some(c => c.toLowerCase() === query)) {
    const rawQ = (item.query || '').trim();
    optionsHtml += `<option value="${escapeHtml(rawQ)}" ${item.selectedValue === rawQ ? 'selected' : ''}>✏️ Use Custom Font: "${escapeHtml(rawQ)}"</option>`;
  }

  // 3. Category grouping when query is empty and category is 'all'
  if (!query && cat === 'all') {
    const pFonts = fontFallbackSystemCatalog.persian_fonts || [];
    const lFonts = fontFallbackSystemCatalog.latin_fonts || [];

    if (pFonts.length > 0) {
      optionsHtml += `<optgroup label="Installed Persian / Multilingual (${pFonts.length})">`;
      pFonts.forEach(f => {
        if (f !== item.rec) {
          optionsHtml += `<option value="${escapeHtml(f)}" ${item.selectedValue === f ? 'selected' : ''}>${escapeHtml(f)}</option>`;
        }
      });
      optionsHtml += `</optgroup>`;
    }

    if (lFonts.length > 0) {
      optionsHtml += `<optgroup label="Installed Latin / Universal (${lFonts.length})">`;
      lFonts.forEach(f => {
        if (f !== item.rec) {
          optionsHtml += `<option value="${escapeHtml(f)}" ${item.selectedValue === f ? 'selected' : ''}>${escapeHtml(f)}</option>`;
        }
      });
      optionsHtml += `</optgroup>`;
    }
  } else {
    // Filtered list by query or specific category
    if (candidates.length === 0) {
      optionsHtml += `<option value="" disabled selected>No matching fonts found for "${escapeHtml(item.query)}"</option>`;
    } else {
      candidates.forEach(f => {
        if (f !== item.rec) {
          optionsHtml += `<option value="${escapeHtml(f)}" ${item.selectedValue === f ? 'selected' : ''}>${escapeHtml(f)}</option>`;
        }
      });
    }
  }

  // Always append Custom option
  optionsHtml += `<option value="__custom__" ${item.selectedValue === '__custom__' ? 'selected' : ''}>✏️ Custom Font Name...</option>`;

  let statsText = '';
  if (query) {
    statsText = `Found ${candidates.length} font(s) matching "${escapeHtml(item.query)}"`;
  } else if (cat === 'persian') {
    statsText = `Showing ${candidates.length} Persian / Arabic fonts`;
  } else if (cat === 'latin') {
    statsText = `Showing ${candidates.length} Latin / Universal fonts`;
  } else {
    statsText = `${candidates.length} installed font families available`;
  }

  return { optionsHtml, statsText };
}

function onFontSearchInput(idx, query) {
  if (!fontFallbackItems[idx]) return;
  fontFallbackItems[idx].query = query;

  const card = document.getElementById(`ff-card-${idx}`);
  if (!card) return;

  const select = document.getElementById(`ff-select-${idx}`);
  const stats = document.getElementById(`ff-stats-${idx}`);
  const curDisplay = document.getElementById(`ff-cur-display-${idx}`);

  const { optionsHtml, statsText } = buildFontFallbackOptions(fontFallbackItems[idx]);

  if (select) {
    select.innerHTML = optionsHtml;
    const queryTrim = (query || '').trim().toLowerCase();
    if (queryTrim) {
      const firstOpt = select.querySelector('option:not([disabled])');
      if (firstOpt) {
        select.value = firstOpt.value;
        fontFallbackItems[idx].selectedValue = firstOpt.value;
        if (curDisplay) curDisplay.innerText = firstOpt.value;
      }
    }
  }
  if (stats) stats.innerText = statsText;

  // Toggle clear button
  const searchInput = document.getElementById(`ff-search-${idx}`);
  let clearBtn = searchInput ? searchInput.parentElement.querySelector('button') : null;
  if (query) {
    if (!clearBtn && searchInput) {
      clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.onclick = () => clearFontSearch(idx);
      clearBtn.className = 'absolute right-2 top-2 text-muted-foreground hover:text-foreground text-xs h-4 w-4 flex items-center justify-center cursor-pointer';
      clearBtn.innerText = '✕';
      clearBtn.title = 'Clear search';
      searchInput.parentElement.appendChild(clearBtn);
    }
  } else if (clearBtn) {
    clearBtn.remove();
  }
}

function onFontSearchEnter(idx) {
  const select = document.getElementById(`ff-select-${idx}`);
  if (!select) return;
  const firstOpt = select.querySelector('option:not([disabled])');
  if (firstOpt) {
    select.value = firstOpt.value;
    onFontFallbackSelectChange(idx, firstOpt.value);
    showToast(`Selected "${firstOpt.value}" for ${fontFallbackItems[idx].missingFont}`, 'info');
  }
}

function clearFontSearch(idx) {
  if (!fontFallbackItems[idx]) return;
  fontFallbackItems[idx].query = '';
  const searchInput = document.getElementById(`ff-search-${idx}`);
  if (searchInput) searchInput.value = '';
  onFontSearchInput(idx, '');
  if (searchInput) searchInput.focus();
}

function setFontCategory(idx, cat) {
  if (!fontFallbackItems[idx]) return;
  fontFallbackItems[idx].category = cat;
  const card = document.getElementById(`ff-card-${idx}`);
  if (card) {
    card.innerHTML = getFontFallbackItemMarkup(fontFallbackItems[idx], idx);
    if (window.lucide) lucide.createIcons();
    const input = document.getElementById(`ff-search-${idx}`);
    if (input) input.focus();
  }
}

function onFontFallbackSelectChange(idx, value) {
  if (!fontFallbackItems[idx]) return;
  const val = value !== undefined ? value : document.getElementById(`ff-select-${idx}`)?.value;
  fontFallbackItems[idx].selectedValue = val;

  const customInput = document.getElementById(`ff-custom-${idx}`);
  const curDisplay = document.getElementById(`ff-cur-display-${idx}`);

  if (val === '__custom__') {
    if (customInput) {
      customInput.classList.remove('hidden');
      customInput.focus();
    }
    if (curDisplay) curDisplay.innerText = fontFallbackItems[idx].customValue || 'Custom';
  } else {
    if (customInput) customInput.classList.add('hidden');
    if (curDisplay) curDisplay.innerText = val;
  }
}

function onFontCustomInput(idx, text) {
  if (!fontFallbackItems[idx]) return;
  fontFallbackItems[idx].customValue = text;
  const curDisplay = document.getElementById(`ff-cur-display-${idx}`);
  if (curDisplay) curDisplay.innerText = text || 'Custom';
}

function applyBatchFontFallback() {
  const batchSelect = document.getElementById('ff-batch-select');
  if (!batchSelect) return;
  const chosenFont = batchSelect.value;
  if (!chosenFont) return;

  fontFallbackItems.forEach(item => {
    item.selectedValue = chosenFont;
    item.query = '';
  });

  renderFontFallbackList();
  showToast(`Applied "${chosenFont}" to all ${fontFallbackItems.length} missing fonts`, 'success');
}

async function submitFontFallbackAction() {
  const fontFallbacks = {};

  fontFallbackItems.forEach(item => {
    let val = item.selectedValue;
    if (val === '__custom__') {
      val = (item.customValue || '').trim() || item.rec || 'Segoe UI';
    }
    fontFallbacks[item.missingFont] = val;
  });

  currentConfiguredFontFallbacks = { ...currentConfiguredFontFallbacks, ...fontFallbacks };

  const btn = document.getElementById('btn-ff-submit');
  if (btn) btn.disabled = true;

  try {
    if (currentHumanTouchJobId) {
      const res = await fetch('/api/generator/human-action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          job_id: currentHumanTouchJobId,
          action: 'continue',
          step: 'font_fallback',
          data: {
            font_fallbacks: fontFallbacks
          }
        })
      });
      const d = await res.json();
      if (!d.success) throw new Error(d.error || 'Failed to apply font fallbacks');

      const mappedStr = Object.entries(fontFallbacks).map(([k, v]) => `${k} → ${v}`).join(', ');
      appendGenLog(`[Step 1.2] User approved font fallbacks: ${mappedStr}`);
      showToast('Font fallbacks applied. Resuming generation...', 'success');
    } else {
      const statusDiv = document.getElementById('gen-template-font-status');
      if (statusDiv) {
        const mappedStr = Object.entries(fontFallbacks).map(([k, v]) => `${k} → ${v}`).join(', ');
        statusDiv.className = 'mt-1.5 text-[11px] p-2 rounded border border-emerald-500/30 bg-emerald-500/10 text-emerald-400 flex items-center justify-between flex-wrap gap-1.5';
        statusDiv.innerHTML = `
          <div class="flex items-center gap-1.5">
            <i data-lucide="check-circle" class="w-3.5 h-3.5 text-emerald-400 shrink-0"></i>
            <span>Fallbacks pre-configured: <strong class="font-mono text-foreground">${escapeHtml(mappedStr)}</strong></span>
          </div>
          <button type="button" onclick="openPreConfigFontFallbackModal()"
            class="shadcn-btn shadcn-btn-outline text-[11px] h-6 px-2 py-0 gap-1 text-muted-foreground hover:text-foreground cursor-pointer">
            Edit
          </button>
        `;
        if (window.lucide) lucide.createIcons();
      }
      showToast('Font fallbacks pre-configured successfully!', 'success');
    }

    closeFontFallbackModal();
  } catch (err) {
    showToast(`Error applying fallbacks: ${err.message}`, 'error');
  } finally {
    if (btn) btn.disabled = false;
  }
}

// -------------------------------------------------------------
// 1.6. AFTER-DONE: AI WORKFLOW PRESENTATION DECK EDITOR
// -------------------------------------------------------------
function setDeckEditPrompt(text) {
  const input = document.getElementById('ht-deck-edit-prompt');
  if (input) {
    input.value = text;
    input.focus();
  }
}

async function submitDeckAiEdit() {
  if (!currentGeneratedPptx) {
    showToast('No presentation generated yet. Generate a deck first.', 'warning');
    return;
  }
  const promptInput = document.getElementById('ht-deck-edit-prompt');
  const prompt = promptInput ? promptInput.value.trim() : '';
  if (!prompt) {
    showToast('Please enter an editing prompt.', 'warning');
    if (promptInput) promptInput.focus();
    return;
  }

  const btnEdit = document.getElementById('btn-ht-deck-edit');
  const statusSpan = document.getElementById('ht-deck-edit-status');
  btnEdit.disabled = true;
  statusSpan.innerText = 'AI Agent modifying presentation...';
  appendGenLog(`\n[Human Touch Deck Edit] Custom prompt: "${prompt}"`);

  try {
    const res = await fetch('/api/generator/edit-pptx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_path: currentGeneratedPptx,
        prompt: prompt
      })
    });
    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || 'Failed to edit presentation.');
    }

    appendGenLog(`[✓] AI Deck Edit completed: ${data.summary} (${data.actions_applied} action(s) applied)`);
    showToast(`Deck updated: ${data.summary}`, 'success');

    if (data.previews && data.previews.length > 0) {
      genSlides = data.previews;
      genSlideIdx = 0;
      visualSlides = [...genSlides];
      visualSlideIdx = 0;
      updateGenSlideDisplay(data.engine_name);
      updateVisualSlideDisplay();
      renderSlideThumbnails();
    }

    if (promptInput) promptInput.value = '';
  } catch (err) {
    appendGenLog(`[!] Deck edit failed: ${err.message}`);
    showToast(`Edit failed: ${err.message}`, 'error');
  } finally {
    btnEdit.disabled = false;
    statusSpan.innerText = '';
  }
}

async function inspectCurrentDeckDetails() {
  if (!currentGeneratedPptx) {
    showToast('No presentation available to inspect.', 'warning');
    return;
  }
  const modal = document.getElementById('inspect-deck-modal');
  const container = document.getElementById('inspect-deck-content');
  if (!modal || !container) return;

  container.innerHTML = '<div class="text-center py-6 text-muted-foreground"><i data-lucide="loader-2" class="w-6 h-6 animate-spin mx-auto mb-2"></i>Inspecting presentation structure...</div>';
  if (window.lucide) lucide.createIcons();
  modal.classList.remove('hidden');

  try {
    const res = await fetch('/api/generator/inspect-deck', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: currentGeneratedPptx })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const deck = data.deck;
    let html = `
      <div class="flex items-center justify-between pb-2 border-b border-border text-xs">
        <span class="font-semibold text-foreground">${deck.filename}</span>
        <span class="text-muted-foreground font-mono">${deck.slide_count} slide(s)</span>
      </div>
    `;

    deck.slides.forEach((s) => {
      html += `
        <div class="p-3 rounded-md bg-secondary/40 border border-border/70 space-y-1.5 text-xs">
          <div class="flex items-center justify-between font-semibold text-foreground">
            <span>Slide ${s.slide_number}: ${s.title || '(Untitled)'}</span>
            <span class="text-[10px] text-muted-foreground font-mono">${s.shapes_count} shapes</span>
          </div>
          ${s.notes ? `<div class="text-[11px] text-muted-foreground italic"><span class="font-medium text-foreground">Notes:</span> ${s.notes}</div>` : ''}
          <div class="space-y-1 pt-1">
            ${s.shapes.map(sh => `
              <div class="text-[11px] bg-background/60 p-1.5 rounded border border-border/40 font-mono flex items-start justify-between gap-2">
                <span class="text-primary shrink-0">#${sh.shape_index} [${sh.type}]</span>
                <span class="text-foreground/90 truncate flex-1 text-right">${sh.text ? sh.text.slice(0, 100) : (sh.has_table ? `Table (${sh.table_data.length} rows)` : (sh.is_picture ? 'Picture' : 'Shape'))}</span>
              </div>
            `).join('')}
          </div>
        </div>
      `;
    });

    container.innerHTML = html;
  } catch (err) {
    container.innerHTML = `<div class="p-4 text-center text-rose-400">Failed to inspect deck: ${err.message}</div>`;
  }
}

function closeInspectDeckModal() {
  const modal = document.getElementById('inspect-deck-modal');
  if (modal) modal.classList.add('hidden');
}

// -------------------------------------------------------------
// 2. TEMPLATE INTELLIGENCE & ANALYZER
// -------------------------------------------------------------
let templateViewMode = 'structured'; // 'structured' | 'editor'

function switchTemplateViewMode(mode) {
  templateViewMode = mode;
  const btnStruct = document.getElementById('btn-view-structured');
  const btnEditor = document.getElementById('btn-view-editor');
  const structView = document.getElementById('tpl-structured-view');
  const editorView = document.getElementById('tpl-editor-view');
  const structActions = document.getElementById('structured-view-actions');
  const editorActions = document.getElementById('editor-actions');

  if (mode === 'structured') {
    btnStruct.classList.add('bg-background', 'text-foreground', 'shadow-sm');
    btnStruct.classList.remove('text-muted-foreground');
    btnEditor.classList.remove('bg-background', 'text-foreground', 'shadow-sm');
    btnEditor.classList.add('text-muted-foreground');

    structView.classList.remove('hidden');
    editorView.classList.add('hidden');
    if (structActions) structActions.classList.remove('hidden');
    if (editorActions) editorActions.classList.add('hidden');

    const tpl = templatesList.find(t => t.filename === selectedTemplateName);
    if (tpl) renderStructuredTemplateView(tpl);
  } else {
    btnEditor.classList.add('bg-background', 'text-foreground', 'shadow-sm');
    btnEditor.classList.remove('text-muted-foreground');
    btnStruct.classList.remove('bg-background', 'text-foreground', 'shadow-sm');
    btnStruct.classList.add('text-muted-foreground');

    structView.classList.add('hidden');
    editorView.classList.remove('hidden');
    if (structActions) structActions.classList.add('hidden');
    if (editorActions) editorActions.classList.remove('hidden');

    loadNoteMd();
  }
}

async function loadTemplatesList() {
  try {
    const res = await fetch('/api/templates/list');
    const data = await res.json();
    templatesList = data.templates || [];

    const structuredCount = templatesList.filter(t => t.is_structured).length;
    document.getElementById('tpl-count-badge').innerText = `${templatesList.length} Templates (${structuredCount} Structured)`;

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

      const statusTag = tpl.is_structured
        ? '<span class="shadcn-badge shadcn-badge-outline text-emerald-400 border-emerald-800/80 bg-emerald-950/20">✓ Structured</span>'
        : (tpl.is_analyzed
          ? '<span class="shadcn-badge shadcn-badge-outline text-sky-400 border-sky-800/80">○ Legacy</span>'
          : '<span class="shadcn-badge shadcn-badge-outline text-muted-foreground">○ Pending</span>');

      const domainTone = tpl.domain && tpl.domain !== 'General'
        ? `${tpl.domain} • ${tpl.style || ''}`
        : (tpl.style || tpl.purpose || 'Standard');

      tr.innerHTML = `
        <td class="py-2 px-3 font-mono text-foreground font-medium">${tpl.filename}</td>
        <td class="py-2 px-3 text-center text-muted-foreground font-mono">${tpl.slide_count}</td>
        <td class="py-2 px-3 text-center">${statusTag}</td>
        <td class="py-2 px-3 text-muted-foreground truncate max-w-[160px]" title="${domainTone}">${domainTone}</td>
      `;
      tbody.appendChild(tr);
    });

    if (!selectedTemplateName && templatesList.length > 0) {
      selectTemplateItem(templatesList[0].filename);
    } else if (selectedTemplateName) {
      const current = templatesList.find(t => t.filename === selectedTemplateName);
      if (current) selectTemplateItem(current.filename);
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
    const isStructBadge = tpl.is_structured
      ? '<span class="shadcn-badge shadcn-badge-outline text-emerald-400 border-emerald-800 text-[10px]">Standardized Schema ✓</span>'
      : '<span class="shadcn-badge shadcn-badge-outline text-amber-400 border-amber-800 text-[10px]">Unstructured Note</span>';

    metaBox.innerHTML = `
      <div class="flex items-start justify-between gap-1 pb-1.5 border-b border-border">
        <div>
          <div class="font-semibold text-foreground text-xs font-mono">${tpl.filename}</div>
          <div class="text-[11px] text-muted-foreground truncate max-w-[170px]">${tpl.display_name || tpl.filename}</div>
        </div>
        ${isStructBadge}
      </div>
      <div class="space-y-1.5 text-[11px] pt-1">
        <div class="flex items-center justify-between text-muted-foreground font-mono">
          <span>Slides: ${tpl.slide_count}</span>
          <span>Dim: ${tpl.dimensions}</span>
        </div>
        <div class="flex flex-wrap gap-1 pt-0.5">
          <span class="px-1.5 py-0.5 rounded bg-secondary text-[10px] text-foreground font-medium">${tpl.density || 'Medium-density'}</span>
          <span class="px-1.5 py-0.5 rounded bg-secondary text-[10px] text-muted-foreground">${tpl.typography || 'Aptos / Arial'}</span>
        </div>
        <div class="text-muted-foreground line-clamp-2 pt-0.5" title="${tpl.purpose}">
          <span class="text-primary font-medium">🎯</span> ${tpl.purpose || 'Not analyzed'}
        </div>
        ${tpl.trigger_conditions ? `<div class="text-[10px] text-muted-foreground/80 line-clamp-1"><span class="text-amber-400 font-medium">Trigger:</span> ${tpl.trigger_conditions}</div>` : ''}
      </div>
    `;

    renderStructuredTemplateView(tpl);
    loadTemplateSlidePreviews(tpl.file_path);
  }
}

function renderStructuredTemplateView(tpl) {
  const container = document.getElementById('tpl-structured-view');
  if (!container) return;

  if (!tpl) {
    container.innerHTML = '<div class="p-6 text-center text-muted-foreground italic">Select a template to view its structured blueprint.</div>';
    return;
  }

  const catalog = tpl.slide_catalog || [];
  let tableRowsHtml = '';

  if (catalog.length > 0) {
    catalog.forEach((sl) => {
      const isSelected = tplSlideIdx === sl.slide_index;
      tableRowsHtml += `
        <tr onclick="jumpToSlidePreview(${sl.slide_index})"
          class="cursor-pointer hover:bg-muted/70 transition ${isSelected ? 'bg-primary/10 border-l-2 border-primary font-medium' : ''}"
          title="Click to preview Slide ${sl.slide_index + 1}">
          <td class="py-1.5 px-2 font-mono text-center text-foreground">${sl.slide_index}</td>
          <td class="py-1.5 px-2 font-mono text-primary text-[11px]">${sl.archetype}</td>
          <td class="py-1.5 px-2 text-foreground truncate max-w-[140px]">${sl.layout_pattern}</td>
          <td class="py-1.5 px-2 text-center font-mono text-muted-foreground">${sl.slot_count}</td>
          <td class="py-1.5 px-2 text-muted-foreground truncate max-w-[180px]">${sl.best_content_fit}</td>
        </tr>
      `;
    });
  } else {
    tableRowsHtml = '<tr><td colspan="5" class="py-3 text-center text-muted-foreground italic">Click "Standardize" to parse slide layout matrix.</td></tr>';
  }

  // Flow Recipe Chips
  let flowRecipeHtml = '';
  if (tpl.flow_recipe && tpl.flow_recipe.length > 0) {
    flowRecipeHtml = tpl.flow_recipe.map((step, i) => `
      <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-primary/10 text-primary border border-primary/20 font-mono text-[11px]">
        <span>S${i + 1}: Slide ${step}</span>
      </span>
    `).join(' <span class="text-muted-foreground">→</span> ');
  } else {
    flowRecipeHtml = '<span class="text-muted-foreground italic">Standard 5-slide flow</span>';
  }

  container.innerHTML = `
    <!-- Top Identity Card -->
    <div class="p-3 bg-secondary/30 rounded-lg border border-border space-y-2">
      <div class="flex items-start justify-between gap-2">
        <div>
          <h3 class="font-semibold text-foreground text-sm flex items-center gap-1.5">
            <i data-lucide="layers" class="w-4 h-4 text-primary"></i> ${tpl.display_name || tpl.filename}
          </h3>
          <p class="text-[11px] text-muted-foreground font-mono mt-0.5">${tpl.filename} • ${tpl.slide_count} Slides • ${tpl.dimensions}</p>
        </div>
        <span class="shadcn-badge shadcn-badge-outline ${tpl.is_structured ? 'text-emerald-400 border-emerald-800' : 'text-amber-400'} font-mono text-[10px]">
          ${tpl.is_structured ? 'Schema v1.0 ✓' : 'Legacy Note'}
        </span>
      </div>

      <!-- Attributes Grid -->
      <div class="grid grid-cols-2 sm:grid-cols-3 gap-2 pt-1">
        <div class="p-1.5 bg-background/60 rounded border border-border/60">
          <div class="text-[10px] text-muted-foreground uppercase font-mono">Domain</div>
          <div class="text-xs text-foreground font-medium truncate" title="${tpl.domain}">${tpl.domain}</div>
        </div>
        <div class="p-1.5 bg-background/60 rounded border border-border/60">
          <div class="text-[10px] text-muted-foreground uppercase font-mono">Tone & Mood</div>
          <div class="text-xs text-foreground font-medium truncate" title="${tpl.style}">${tpl.style}</div>
        </div>
        <div class="p-1.5 bg-background/60 rounded border border-border/60">
          <div class="text-[10px] text-muted-foreground uppercase font-mono">Density</div>
          <div class="text-xs text-foreground font-medium truncate">${tpl.density || 'Medium'}</div>
        </div>
        <div class="p-1.5 bg-background/60 rounded border border-border/60">
          <div class="text-[10px] text-muted-foreground uppercase font-mono">Theme / Color</div>
          <div class="text-xs text-foreground font-medium truncate" title="${tpl.color_theme}">${tpl.color_theme}</div>
        </div>
        <div class="p-1.5 bg-background/60 rounded border border-border/60">
          <div class="text-[10px] text-muted-foreground uppercase font-mono">Typography</div>
          <div class="text-xs text-foreground font-medium truncate font-mono">${tpl.typography}</div>
        </div>
        <div class="p-1.5 bg-background/60 rounded border border-border/60">
          <div class="text-[10px] text-muted-foreground uppercase font-mono">Slide Ratio</div>
          <div class="text-xs text-foreground font-medium font-mono">${tpl.dimensions}</div>
        </div>
      </div>
    </div>

    <!-- Executive Purpose & Ideal Use Cases -->
    <div class="p-3 bg-secondary/20 rounded-lg border border-border space-y-2">
      <div class="font-semibold text-foreground text-xs flex items-center gap-1.5">
        <i data-lucide="target" class="w-3.5 h-3.5 text-primary"></i> Executive Purpose & AI Selection Target
      </div>
      <p class="text-muted-foreground text-xs leading-relaxed italic bg-background/40 p-2 rounded border border-border/40">
        "${tpl.purpose}"
      </p>
      ${tpl.trigger_conditions ? `
        <div class="text-[11px] text-muted-foreground flex items-center gap-1.5">
          <span class="text-amber-400 font-semibold">Trigger Condition:</span>
          <span>${tpl.trigger_conditions}</span>
        </div>
      ` : ''}
    </div>

    <!-- Master Slide Architecture & Slot Matrix -->
    <div class="space-y-1.5">
      <div class="flex items-center justify-between">
        <div class="font-semibold text-foreground text-xs flex items-center gap-1.5">
          <i data-lucide="layout-template" class="w-3.5 h-3.5 text-primary"></i> Master Slide Architecture & Slot Matrix
        </div>
        <span class="text-[10px] text-muted-foreground font-mono">Click row to preview slide</span>
      </div>
      <div class="overflow-x-auto rounded-md border border-border max-h-[220px] overflow-y-auto">
        <table class="w-full text-left text-xs">
          <thead class="bg-muted text-[10px] uppercase font-mono text-muted-foreground sticky top-0 border-b border-border">
            <tr>
              <th class="py-1.5 px-2 text-center w-10">Slide</th>
              <th class="py-1.5 px-2">Archetype</th>
              <th class="py-1.5 px-2">Layout Pattern</th>
              <th class="py-1.5 px-2 text-center w-12">Slots</th>
              <th class="py-1.5 px-2">Best Content Fit</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-border/60 font-mono text-[11px]">
            ${tableRowsHtml}
          </tbody>
        </table>
      </div>
    </div>

    <!-- AI Directives & Flow Recipes -->
    <div class="p-3 bg-secondary/20 rounded-lg border border-border space-y-2">
      <div class="font-semibold text-foreground text-xs flex items-center gap-1.5">
        <i data-lucide="git-merge" class="w-3.5 h-3.5 text-sky-400"></i> AI Generator Directives & Sequencing Recipes
      </div>
      <div class="space-y-1.5 text-[11px]">
        <div class="flex items-center gap-2 flex-wrap">
          <span class="text-muted-foreground font-mono">Recommended 5-Slide Flow:</span>
          <div class="flex items-center gap-1 flex-wrap">${flowRecipeHtml}</div>
        </div>
        <div class="text-[11px] text-muted-foreground flex items-center gap-1">
          <span class="text-emerald-400 font-medium">✓ Slot Budget Protection:</span>
          <span>Enforces strict character limits per slot to eliminate text overflow and clipping.</span>
        </div>
        <div class="text-[11px] text-muted-foreground flex items-center gap-1">
          <span class="text-amber-400 font-medium">✂ Shape Pruning Strategy:</span>
          <span>Unused column cards and milestone badges are added to shapes_to_remove.</span>
        </div>
      </div>
    </div>
  `;

  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function jumpToSlidePreview(slideIdx) {
  if (tplSlides && slideIdx >= 0 && slideIdx < tplSlides.length) {
    tplSlideIdx = slideIdx;
    updateTplSlideDisplay();
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
  const badge = document.getElementById('tpl-slide-arch-badge');

  ph.innerText = 'Rendering slides...';
  ph.classList.remove('hidden');
  img.classList.add('hidden');
  if (badge) badge.classList.add('hidden');

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
  const badge = document.getElementById('tpl-slide-arch-badge');

  if (!tplSlides || tplSlides.length === 0) {
    img.classList.add('hidden');
    ph.classList.remove('hidden');
    if (badge) badge.classList.add('hidden');
    counter.innerText = '0 / 0';
    return;
  }

  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = tplSlides[tplSlideIdx].data_url;
  counter.innerText = `${tplSlideIdx + 1} / ${tplSlides.length}`;

  // Update Archetype Badge if available
  const currentTpl = templatesList.find(t => t.filename === selectedTemplateName);
  if (badge && currentTpl && currentTpl.slide_catalog && currentTpl.slide_catalog[tplSlideIdx]) {
    const slInfo = currentTpl.slide_catalog[tplSlideIdx];
    badge.innerText = `${slInfo.archetype} • ${slInfo.layout_pattern}`;
    badge.classList.remove('hidden');
  } else if (badge) {
    badge.classList.add('hidden');
  }
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

async function applyStandardSchema(applyAll = false) {
  if (applyAll) {
    if (!confirm('Apply Standard 4-Part Schema to ALL templates in data/*.pptx? This will standardize data/NOTE.md.')) {
      return;
    }
  }

  const targetFilename = applyAll ? null : selectedTemplateName;
  if (!applyAll && !targetFilename) {
    showToast('Please select a template from the list first.', 'warning');
    return;
  }

  showToast(`Formatting ${applyAll ? 'all templates' : targetFilename} into standard structured schema...`, 'info');
  try {
    const res = await fetch('/api/templates/apply-schema', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: targetFilename, all: applyAll })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    showToast(data.message || 'Successfully applied standard schema!', 'success');
    await loadTemplatesList();
    await loadNoteMd();
  } catch (err) {
    showToast(`Standardize failed: ${err.message}`, 'error');
  }
}

async function openTemplateSchemaModal() {
  const modal = document.getElementById('template-schema-modal');
  const codeBox = document.getElementById('schema-spec-code');
  if (!modal) return;

  modal.classList.remove('hidden');
  codeBox.innerText = 'Loading schema specification...';

  try {
    const res = await fetch('/api/templates/schema');
    const data = await res.json();
    if (data.success && data.schema) {
      codeBox.innerText = data.schema;
    } else {
      codeBox.innerText = 'Could not load schema specification.';
    }
  } catch (err) {
    codeBox.innerText = `Error: ${err.message}`;
  }

  if (window.lucide) window.lucide.createIcons();
}

function closeTemplateSchemaModal() {
  const modal = document.getElementById('template-schema-modal');
  if (modal) modal.classList.add('hidden');
}

function copySchemaSpecToClipboard() {
  const codeBox = document.getElementById('schema-spec-code');
  if (!codeBox) return;
  navigator.clipboard.writeText(codeBox.innerText)
    .then(() => showToast('Schema specification copied to clipboard!', 'success'))
    .catch(() => showToast('Failed to copy to clipboard', 'error'));
}

async function insertSchemaBoilerplate() {
  try {
    const filename = selectedTemplateName || 'template.pptx';
    const res = await fetch(`/api/templates/boilerplate?filename=${encodeURIComponent(filename)}`);
    const data = await res.json();
    if (data.success && data.boilerplate) {
      const editor = document.getElementById('note-md-editor');
      if (editor) {
        editor.value = (editor.value.trim() ? editor.value.trim() + '\n\n---\n\n' : '') + data.boilerplate;
        showToast(`Inserted structured schema boilerplate for ${filename}`, 'success');
      }
    }
  } catch (err) {
    showToast(`Failed to load boilerplate: ${err.message}`, 'error');
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
  if (btnSel) btnSel.disabled = true;
  if (btnAll) btnAll.disabled = true;

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
      if (btnSel) btnSel.disabled = false;
      if (btnAll) btnAll.disabled = false;
      showToast(`Analysis completed for ${selectedTemplateName}`, 'success');
    });
    evtSource.addEventListener('error', () => {
      appendAnalyzeLog(`[!] Error analyzing template`);
      if (btnSel) btnSel.disabled = false;
      if (btnAll) btnAll.disabled = false;
      showToast('Template analysis error', 'error');
    });
    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAnalyzeLog(`[!] Analysis failed: ${err.message}`);
    if (btnSel) btnSel.disabled = false;
    if (btnAll) btnAll.disabled = false;
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

  if (btnSel) btnSel.disabled = true;
  if (btnAll) btnAll.disabled = true;
  if (progContainer) progContainer.classList.remove('hidden');
  if (progBar) progBar.style.width = '0%';

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
      if (progBar) progBar.style.width = `${d.percentage}%`;
      if (progPct) progPct.innerText = `${d.percentage}%`;
      if (progStatus) progStatus.innerText = `Analyzing [${d.current}/${d.total}]: ${d.current_name}`;
    });

    evtSource.addEventListener('completed', () => {
      appendAnalyzeLog(`[✓] Completed batch analysis of all templates.`);
      if (progBar) progBar.style.width = '100%';
      if (progPct) progPct.innerText = '100%';
      loadTemplatesList();
      loadNoteMd();
      if (btnSel) btnSel.disabled = false;
      if (btnAll) btnAll.disabled = false;
      showToast('Batch template analysis complete!', 'success');
    });

    evtSource.addEventListener('error', () => {
      appendAnalyzeLog(`[!] Error in batch template analysis.`);
      if (btnSel) btnSel.disabled = false;
      if (btnAll) btnAll.disabled = false;
      showToast('Batch analysis failed', 'error');
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAnalyzeLog(`[!] Batch analysis failed: ${err.message}`);
    if (btnSel) btnSel.disabled = false;
    if (btnAll) btnAll.disabled = false;
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

let currentMgrEngine = '';

function updateMgrSlideDisplay(engineName = '') {
  const ph = document.getElementById('mgr-slide-placeholder');
  const img = document.getElementById('mgr-slide-img');
  const counter = document.getElementById('mgr-slide-counter');
  const prevBtn = document.getElementById('mgr-prev-btn');
  const nextBtn = document.getElementById('mgr-next-btn');
  const badge = document.getElementById('mgr-engine-badge');

  if (engineName) {
    currentMgrEngine = engineName;
  }

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

  const activeEngine = currentMgrEngine || 'Native PowerPoint';
  if (badge) {
    badge.classList.remove('hidden');
    if (activeEngine.includes('PowerPoint')) {
      badge.innerText = 'Native PowerPoint';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 border border-emerald-800 text-emerald-300';
    } else if (activeEngine.includes('Web')) {
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
// 6. SETTINGS & CONFIGURATION (9ROUTER MODEL DISCOVERY & SUGGESTIONS)
// -------------------------------------------------------------
let _9routerModelsData = null;
let _selectedProviderFilter = 'all';
let _modelBrowserOpen = false;

function getProviderBadgeClass(owner) {
  const o = (owner || '').toLowerCase();
  if (o === 'combo') return 'bg-purple-500/15 text-purple-400 border-purple-500/30';
  if (o === 'ag') return 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30';
  if (o === 'gemini') return 'bg-blue-500/15 text-blue-400 border-blue-500/30';
  if (o === 'openrouter') return 'bg-amber-500/15 text-amber-400 border-amber-500/30';
  if (o === 'openai') return 'bg-teal-500/15 text-teal-400 border-teal-500/30';
  if (o === 'anthropic') return 'bg-orange-500/15 text-orange-400 border-orange-500/30';
  if (o === 'deepseek') return 'bg-cyan-500/15 text-cyan-400 border-cyan-500/30';
  if (o === 'aval') return 'bg-indigo-500/15 text-indigo-400 border-indigo-500/30';
  return 'bg-secondary text-foreground border-border';
}

function formatTokensFriendly(num) {
  if (!num) return '';
  if (num >= 1000000) {
    return `${(num / 1000000).toFixed(1).replace('.0', '')}M`;
  }
  return `${Math.round(num / 1024)}k`;
}

async function fetch9RouterModels(forceRefresh = false) {
  const statusEl = document.getElementById('model-browser-status');
  const badgeEl = document.getElementById('model-count-badge');
  const dotEl = document.getElementById('model-status-dot');
  const refreshIcon = document.getElementById('model-refresh-icon');

  if (refreshIcon) refreshIcon.classList.add('animate-spin');
  if (statusEl) statusEl.innerText = '(fetching models...)';

  try {
    const rawUrl = document.getElementById('cfg-url')?.value.trim();
    const rawKey = document.getElementById('cfg-key')?.value.trim();
    const params = new URLSearchParams({ category: 'all' });
    if (rawUrl) params.append('url', rawUrl);
    if (rawKey) params.append('key', rawKey);
    if (forceRefresh) params.append('refresh', '1');

    const res = await fetch(`/api/models?${params.toString()}`);
    const data = await res.json();
    if (data && data.success) {
      _9routerModelsData = data;

      const chatCount = data.categories?.chat?.length || data.models?.length || 0;
      if (badgeEl) badgeEl.innerText = `${chatCount}`;
      if (statusEl) {
        statusEl.innerText = data.connected
          ? `(${chatCount} models live from 9Router)`
          : `(${chatCount} cached/fallback models)`;
      }
      if (dotEl) {
        dotEl.className = data.connected
          ? 'w-2 h-2 rounded-full bg-emerald-500 animate-pulse'
          : 'w-2 h-2 rounded-full bg-amber-500';
      }

      // Populate Datalists for instant autocomplete
      populateModelDatalists(data);

      // Populate Quick Suggestion Chips
      populateQuickModelChips(data);

      // Render Provider Filter Tabs
      renderProviderTabs(data.providers || []);

      // Render Model Cards in Browser
      renderModelCards();

      if (forceRefresh) {
        showToast(`Discovered ${chatCount} models from 9Router`, 'success', 2500);
      }
    }
  } catch (err) {
    console.error('Failed to fetch 9Router models:', err);
    if (statusEl) statusEl.innerText = '(offline fallback)';
    if (dotEl) dotEl.className = 'w-2 h-2 rounded-full bg-destructive';
  } finally {
    if (refreshIcon) refreshIcon.classList.remove('animate-spin');
  }
}

function populateModelDatalists(data) {
  const cats = data.categories || {};

  // Chat models datalist
  const chatDl = document.getElementById('cfg-chat-model-datalist');
  if (chatDl && cats.chat) {
    chatDl.innerHTML = cats.chat.map(m =>
      `<option value="${escapeHtml(m.id)}">${escapeHtml(m.provider || m.owned_by)}${m.context_length ? ` • Context: ${formatTokensFriendly(m.context_length)}` : ''}</option>`
    ).join('');
  }

  // Search models datalist
  const searchDl = document.getElementById('cfg-search-model-datalist');
  if (searchDl && cats.search) {
    searchDl.innerHTML = cats.search.map(m =>
      `<option value="${escapeHtml(m.id)}">${escapeHtml(m.provider || m.owned_by)}</option>`
    ).join('');
  }

  // Fetch models datalist
  const fetchDl = document.getElementById('cfg-fetch-model-datalist');
  if (fetchDl && cats.fetch) {
    fetchDl.innerHTML = cats.fetch.map(m =>
      `<option value="${escapeHtml(m.id)}">${escapeHtml(m.provider || m.owned_by)}</option>`
    ).join('');
  }

  // Image models datalist
  const imgDl = document.getElementById('cfg-image-model-datalist');
  if (imgDl && cats.image) {
    imgDl.innerHTML = cats.image.map(m =>
      `<option value="${escapeHtml(m.id)}">${escapeHtml(m.provider || m.owned_by)}</option>`
    ).join('');
  }
}

function populateQuickModelChips(data) {
  const currentChatModel = document.getElementById('cfg-chat-model')?.value.trim();
  const cats = data.categories || {};

  // Chat chips
  const chatChipsEl = document.getElementById('quick-model-chips');
  if (chatChipsEl) {
    const recs = data.recommended || [
      'ag/gemini-3.8-flash-low',
      'ag/gemini-3.8-flash-high',
      'light-code',
      'claude-pro-agent',
      'openai/gpt-4o',
      'aval/deepseek-v4.1-flash'
    ];
    chatChipsEl.innerHTML = recs.map(id => {
      const isSelected = (currentChatModel === id);
      return `<button type="button" onclick="selectSuggestedModel('${escapeHtml(id)}')"
        class="shadcn-badge cursor-pointer px-2 py-0.5 rounded-full font-mono text-[11px] transition-all ${
          isSelected
            ? 'bg-primary text-primary-foreground border-primary font-bold shadow-sm'
            : 'shadcn-badge-outline hover:border-primary hover:text-primary'
        }">
        ${escapeHtml(id)}
      </button>`;
    }).join('');
  }

  // Search chips
  const searchChipsEl = document.getElementById('quick-search-chips');
  if (searchChipsEl) {
    const list = ['exa/search', 'tavily', 'brave-search'];
    searchChipsEl.innerHTML = list.map(id =>
      `<button type="button" onclick="document.getElementById('cfg-search-model').value='${id}'; showToast('Selected search model: ${id}', 'info', 1500);"
        class="shadcn-badge shadcn-badge-outline cursor-pointer px-2 py-0.5 rounded-full font-mono text-[10px] hover:border-primary hover:text-primary">
        ${id}
      </button>`
    ).join('');
  }

  // Fetch chips
  const fetchChipsEl = document.getElementById('quick-fetch-chips');
  if (fetchChipsEl) {
    const list = ['exa/fetch', 'jina-reader', 'firecrawl'];
    fetchChipsEl.innerHTML = list.map(id =>
      `<button type="button" onclick="document.getElementById('cfg-fetch-model').value='${id}'; showToast('Selected fetch model: ${id}', 'info', 1500);"
        class="shadcn-badge shadcn-badge-outline cursor-pointer px-2 py-0.5 rounded-full font-mono text-[10px] hover:border-primary hover:text-primary">
        ${id}
      </button>`
    ).join('');
  }

  // Image chips
  const imgChipsEl = document.getElementById('quick-image-chips');
  if (imgChipsEl) {
    const list = (cats.image && cats.image.length > 0)
      ? cats.image.map(m => m.id)
      : ['gemini/gemini-3.1-flash-image-preview', 'gemini/gemini-3-pro-image-preview', 'ag/gemini-3.1-flash-image'];
    imgChipsEl.innerHTML = list.map(id =>
      `<button type="button" onclick="document.getElementById('cfg-image-model').value='${id}'; showToast('Selected image model: ${id}', 'info', 1500);"
        class="shadcn-badge shadcn-badge-outline cursor-pointer px-2 py-0.5 rounded-full font-mono text-[10px] hover:border-primary hover:text-primary">
        ${id}
      </button>`
    ).join('');
  }
}

function renderProviderTabs(providers) {
  const container = document.getElementById('model-provider-tabs');
  if (!container) return;

  const cats = _9routerModelsData?.categories?.chat || [];
  const counts = { all: cats.length };
  cats.forEach(m => {
    const o = m.owned_by || 'other';
    counts[o] = (counts[o] || 0) + 1;
  });

  const allProviders = ['all', ...providers];
  container.innerHTML = allProviders.map(p => {
    const isSelected = (_selectedProviderFilter === p);
    const count = counts[p] || 0;
    const label = p === 'all' ? 'All' : (p === 'combo' ? 'Combos' : p.toUpperCase());
    return `<button type="button" onclick="setProviderFilter('${escapeHtml(p)}')"
      class="px-2 py-0.5 rounded text-[11px] font-mono transition-all cursor-pointer ${
        isSelected
          ? 'bg-primary text-primary-foreground font-semibold shadow-xs'
          : 'bg-background/80 hover:bg-background text-muted-foreground hover:text-foreground border border-border/80'
      }">
      ${escapeHtml(label)} <span class="opacity-70 text-[10px]">(${count})</span>
    </button>`;
  }).join('');
}

function toggleModelBrowser(forceOpen = null) {
  const drawer = document.getElementById('model-browser-drawer');
  const chevron = document.getElementById('model-browser-chevron');
  if (!drawer) return;

  _modelBrowserOpen = (forceOpen !== null) ? forceOpen : drawer.classList.contains('hidden');
  if (_modelBrowserOpen) {
    drawer.classList.remove('hidden');
    if (chevron) chevron.classList.add('rotate-180');
    if (!_9routerModelsData) {
      fetch9RouterModels();
    } else {
      renderModelCards();
    }
    setTimeout(() => {
      document.getElementById('model-search-input')?.focus();
    }, 50);
  } else {
    drawer.classList.add('hidden');
    if (chevron) chevron.classList.remove('rotate-180');
  }
  if (window.lucide) lucide.createIcons();
}

function setProviderFilter(provider) {
  _selectedProviderFilter = provider;
  if (_9routerModelsData?.providers) {
    renderProviderTabs(_9routerModelsData.providers);
  }
  renderModelCards();
}

function filterModelsList() {
  renderModelCards();
}

function renderModelCards() {
  const container = document.getElementById('model-cards-container');
  if (!container || !_9routerModelsData) return;

  const chatModels = _9routerModelsData.categories?.chat || _9routerModelsData.models || [];
  const currentChatModel = document.getElementById('cfg-chat-model')?.value.trim();
  const search = (document.getElementById('model-search-input')?.value || '').trim().toLowerCase();

  let filtered = chatModels.filter(m => {
    // Provider filter
    if (_selectedProviderFilter !== 'all' && m.owned_by !== _selectedProviderFilter) {
      return false;
    }
    // Search query filter
    if (search) {
      const matchId = (m.id || '').toLowerCase().includes(search);
      const matchProvider = (m.provider || '').toLowerCase().includes(search);
      const matchOwner = (m.owned_by || '').toLowerCase().includes(search);
      if (!matchId && !matchProvider && !matchOwner) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="p-6 text-center text-muted-foreground font-mono text-xs">
        <i data-lucide="search-x" class="w-6 h-6 mx-auto mb-1.5 opacity-40"></i>
        No 9Router models match "${escapeHtml(search)}".
      </div>
    `;
    if (window.lucide) lucide.createIcons();
    return;
  }

  container.innerHTML = filtered.map(m => {
    const isSelected = (currentChatModel === m.id);
    const badgeClass = getProviderBadgeClass(m.owned_by);

    let capsHtml = '';
    if (m.capabilities?.vision) capsHtml += '<span class="text-emerald-400">👁 Vision</span>';
    if (m.capabilities?.tools) capsHtml += '<span class="text-cyan-400">⚡ Tools</span>';
    if (m.capabilities?.reasoning) capsHtml += '<span class="text-amber-400">🧠 Reasoning</span>';
    if (m.capabilities?.search) capsHtml += '<span class="text-sky-400">🔍 Search</span>';

    return `
      <div class="group flex items-center justify-between p-2.5 rounded-md hover:bg-secondary/70 border ${
        isSelected ? 'border-primary/60 bg-primary/5' : 'border-border/40 bg-background/50'
      } transition-all cursor-pointer" onclick="selectSuggestedModel('${escapeHtml(m.id)}')">
        <div class="space-y-1 min-w-0 pr-2">
          <div class="flex items-center gap-1.5 flex-wrap">
            <span class="font-mono text-xs font-semibold ${isSelected ? 'text-primary' : 'text-foreground'} truncate">
              ${escapeHtml(m.id)}
            </span>
            <span class="text-[10px] px-1.5 py-0.2 rounded border font-mono ${badgeClass}">
              ${escapeHtml(m.provider || m.owned_by)}
            </span>
            ${m.is_combo ? '<span class="text-[10px] px-1.5 py-0.2 rounded bg-purple-950/40 border border-purple-700/50 text-purple-300 font-mono">Combo Auto-Fallback</span>' : ''}
          </div>
          <div class="flex items-center gap-2.5 text-[10px] text-muted-foreground font-mono flex-wrap">
            ${m.context_length ? `<span>Ctx: ${formatTokensFriendly(m.context_length)}</span>` : ''}
            ${m.max_tokens ? `<span>Max: ${formatTokensFriendly(m.max_tokens)}</span>` : ''}
            ${capsHtml}
          </div>
        </div>
        <div class="flex items-center gap-2 shrink-0">
          <button type="button" class="text-xs px-2.5 py-1 rounded transition-all ${
            isSelected
              ? 'bg-primary text-primary-foreground font-semibold shadow-xs'
              : 'shadcn-btn-outline hover:bg-primary/20 hover:text-primary'
          }">
            ${isSelected ? '✓ Active' : 'Select'}
          </button>
        </div>
      </div>
    `;
  }).join('');

  if (window.lucide) lucide.createIcons();
}

// -------------------------------------------------------------
// DEDICATED PER-AGENT MODELS & REASONING CONFIGURATION
// -------------------------------------------------------------
const AGENT_IDS = ['generator', 'verifier', 'autonomous', 'analyzer', 'structure', 'editor', 'parser'];
let _targetAgentForModelPick = null;

function pickModelForAgent(agentId) {
  _targetAgentForModelPick = agentId;
  const targetNameEl = document.getElementById('agent-picker-target-name');
  if (targetNameEl) {
    const titles = {
      generator: 'Slide Synthesis & Blueprint',
      verifier: 'Slide Verification & QA',
      autonomous: 'Autonomous AI Assistant',
      analyzer: 'Template Analyzer',
      structure: 'Storyboard & Structure',
      editor: 'Human Touch / Slide Editor',
      parser: 'Document Parser & OCR'
    };
    targetNameEl.innerText = titles[agentId] || agentId.toUpperCase();
  }
  const banner = document.getElementById('agent-model-picker-banner');
  if (banner) banner.classList.remove('hidden');
  toggleModelBrowser(true);
  showToast(`Select a 9Router model for ${agentId} agent from the list`, 'info', 2500);
}

function cancelAgentModelPick() {
  _targetAgentForModelPick = null;
  const banner = document.getElementById('agent-model-picker-banner');
  if (banner) banner.classList.add('hidden');
}

function updateAgentStatusBadge(agentId, primaryModel = null) {
  const modelInput = document.getElementById(`cfg-agent-model-${agentId}`);
  const thinkSelect = document.getElementById(`cfg-agent-think-${agentId}`);
  const statusBadge = document.getElementById(`cfg-agent-status-${agentId}`);
  if (!statusBadge) return;

  const currentPrimary = primaryModel || document.getElementById('cfg-chat-model')?.value.trim() || 'Default';
  const customModel = (modelInput?.value || '').trim();
  const thinkVal = (thinkSelect?.value || 'default').toLowerCase();

  const thinkLabels = {
    default: 'Default',
    none: 'Off',
    low: 'Low',
    medium: 'Medium',
    high: 'High'
  };
  const thinkText = thinkLabels[thinkVal] || thinkVal.toUpperCase();

  if (modelInput && !modelInput.value.trim()) {
    const shortPrimary = currentPrimary.split('/').pop() || currentPrimary;
    modelInput.placeholder = `Inherit (${shortPrimary})`;
  }

  if (customModel && customModel.toLowerCase() !== 'default') {
    const shortName = customModel.split('/').pop() || customModel;
    statusBadge.innerText = `Custom: ${shortName} • ${thinkText}`;
    statusBadge.title = `Model: ${customModel} | Thinking: ${thinkText}`;
    statusBadge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-primary/20 text-primary border border-primary/30';
  } else {
    statusBadge.innerText = `Inherited • ${thinkText}`;
    statusBadge.title = `Inheriting primary: ${currentPrimary} | Thinking: ${thinkText}`;
    statusBadge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-secondary text-muted-foreground border border-border/40';
  }
}

function updateAgentCustomizedSummary() {
  const summaryEl = document.getElementById('agent-customized-summary');
  if (!summaryEl) return;
  let customCount = 0;
  AGENT_IDS.forEach(aid => {
    const val = document.getElementById(`cfg-agent-model-${aid}`)?.value.trim();
    const thk = document.getElementById(`cfg-agent-think-${aid}`)?.value;
    if (val || (thk && thk !== 'default')) {
      customCount++;
    }
  });
  if (customCount === 0) {
    summaryEl.innerText = 'All 7 agents currently inherit the primary chat model';
  } else {
    summaryEl.innerText = `${customCount} of 7 agents customized with dedicated settings`;
  }
}

function onAgentConfigChanged(agentId) {
  updateAgentStatusBadge(agentId);
  updateAgentCustomizedSummary();
}

function resetAgentToDefault(agentId) {
  const modelInput = document.getElementById(`cfg-agent-model-${agentId}`);
  if (modelInput) modelInput.value = '';
  const thinkSelect = document.getElementById(`cfg-agent-think-${agentId}`);
  if (thinkSelect) thinkSelect.value = 'default';
  updateAgentStatusBadge(agentId);
  updateAgentCustomizedSummary();
  showToast(`Reset ${agentId} agent to inherit primary model`, 'info', 1500);
}

function resetAllAgentsToDefault() {
  AGENT_IDS.forEach(aid => {
    const modelInput = document.getElementById(`cfg-agent-model-${aid}`);
    if (modelInput) modelInput.value = '';
    const thinkSelect = document.getElementById(`cfg-agent-think-${aid}`);
    if (thinkSelect) thinkSelect.value = 'default';
    updateAgentStatusBadge(aid);
  });
  updateAgentCustomizedSummary();
  showToast('All agents reset to inherit primary model and default thinking', 'info', 1800);
}

function applyAgentThinkingPreset(level) {
  AGENT_IDS.forEach(aid => {
    const thinkSelect = document.getElementById(`cfg-agent-think-${aid}`);
    if (thinkSelect) thinkSelect.value = level;
    updateAgentStatusBadge(aid);
  });
  updateAgentCustomizedSummary();
  showToast(`Applied ${level.toUpperCase()} thinking level to all agents`, 'success', 2000);
}

function updateAllAgentStatusBadges(data = null) {
  const primaryModel = data?.config?.NINEROUTER_CHAT_MODEL || document.getElementById('cfg-chat-model')?.value.trim() || 'Default';
  AGENT_IDS.forEach(aid => {
    updateAgentStatusBadge(aid, primaryModel);
  });
  updateAgentCustomizedSummary();
}

async function selectSuggestedModel(modelId) {
  // If an agent model browse action was triggered, assign model to that agent!
  if (_targetAgentForModelPick) {
    const aid = _targetAgentForModelPick;
    const input = document.getElementById(`cfg-agent-model-${aid}`);
    if (input) {
      input.value = modelId;
    }
    cancelAgentModelPick();
    onAgentConfigChanged(aid);
    showToast(`Assigned ${modelId} to ${aid} agent`, 'success', 2500);
    await saveConfigSettings();
    return;
  }

  const input = document.getElementById('cfg-chat-model');
  if (input) {
    input.value = modelId;
  }

  // Update quick chips active state
  if (_9routerModelsData) {
    populateQuickModelChips(_9routerModelsData);
    renderModelCards();
  }

  // Update Model Intelligence metadata view immediately if model data exists
  const chatModels = _9routerModelsData?.categories?.chat || [];
  const found = chatModels.find(m => m.id === modelId);
  if (found) {
    const maxTokens = found.max_tokens ? Number(found.max_tokens).toLocaleString() : '--';
    const contextLength = found.context_length ? Number(found.context_length).toLocaleString() : '--';

    const elMax = document.getElementById('cfg-meta-max-tokens');
    if (elMax) elMax.innerText = `${maxTokens} tokens`;

    const elCtx = document.getElementById('cfg-meta-context');
    if (elCtx) elCtx.innerText = `${contextLength} tokens`;

    const elCaps = document.getElementById('cfg-meta-caps');
    if (elCaps && found.capabilities) {
      elCaps.innerHTML = '';
      const capList = [
        { key: 'vision', label: 'Vision OCR' },
        { key: 'tools', label: 'Tool Calling' },
        { key: 'reasoning', label: 'Extended Reasoning' },
        { key: 'search', label: 'Web Search' }
      ];
      capList.forEach(c => {
        const active = Boolean(found.capabilities[c.key]);
        const span = document.createElement('span');
        span.className = `shadcn-badge shadcn-badge-outline text-[10px] py-0.5 px-1.5 font-mono ${active ? 'text-emerald-400 border-emerald-800/80 bg-emerald-950/30' : 'text-muted-foreground border-border'}`;
        span.innerText = `${active ? '✓' : '○'} ${c.label}`;
        elCaps.appendChild(span);
      });
    }
  }

  // Auto-save and apply to backend & .env immediately
  await saveConfigSettings();
}

function onModelInputChanged(val) {
  const trimmed = (val || '').trim();
  if (!_9routerModelsData) return;

  // Highlight quick chip if matches
  populateQuickModelChips(_9routerModelsData);

  // Look up model and live-update intelligence preview
  const chatModels = _9routerModelsData.categories?.chat || [];
  const found = chatModels.find(m => m.id.toLowerCase() === trimmed.toLowerCase());
  if (found) {
    const maxTokens = found.max_tokens ? Number(found.max_tokens).toLocaleString() : '--';
    const contextLength = found.context_length ? Number(found.context_length).toLocaleString() : '--';
    const elMax = document.getElementById('cfg-meta-max-tokens');
    if (elMax) elMax.innerText = `${maxTokens} tokens`;
    const elCtx = document.getElementById('cfg-meta-context');
    if (elCtx) elCtx.innerText = `${contextLength} tokens`;
  }
}

async function refresh9RouterModels(forceLive = false) {
  showToast('Connecting to 9Router gateway...', 'info', 1200);
  await fetch9RouterModels(forceLive);
}

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
      const vRoundsEl = document.getElementById('cfg-verification-rounds');
      if (vRoundsEl) vRoundsEl.value = cfg.VERIFICATION_ROUNDS || 3;
      const genVRoundsEl = document.getElementById('gen-verification-rounds');
      if (genVRoundsEl) genVRoundsEl.value = cfg.VERIFICATION_ROUNDS || 3;

      // Populate per-agent model and think level configurations
      AGENT_IDS.forEach(aid => {
        const mKey = 'AGENT_MODEL_' + aid.toUpperCase();
        const tKey = 'AGENT_THINK_LEVEL_' + aid.toUpperCase();
        const mInput = document.getElementById(`cfg-agent-model-${aid}`);
        const tSelect = document.getElementById(`cfg-agent-think-${aid}`);
        if (mInput) {
          mInput.value = cfg[mKey] || (data.agents && data.agents[aid]?.configured_model) || '';
        }
        if (tSelect) {
          tSelect.value = cfg[tKey] || (data.agents && data.agents[aid]?.think_level) || 'default';
        }
      });

      applyConfigAndMetadata(data);
      updateAllAgentStatusBadges(data);
      // Automatically load 9Router model suggestions & datalists
      fetch9RouterModels();
    }
  } catch (err) {
    console.error('Failed to load settings', err);
  }
}

async function saveConfigSettings() {
  const timeoutVal = parseInt(document.getElementById('cfg-timeout')?.value.trim() || '300', 10);
  const vRoundsVal = parseInt(document.getElementById('cfg-verification-rounds')?.value.trim() || '3', 10);
  const chatModel = document.getElementById('cfg-chat-model')?.value.trim();

  const config = {
    NINEROUTER_URL: document.getElementById('cfg-url')?.value.trim() || '',
    NINEROUTER_KEY: document.getElementById('cfg-key')?.value.trim() || '',
    NINEROUTER_CHAT_MODEL: chatModel,
    NINEROUTER_SEARCH_MODEL: document.getElementById('cfg-search-model')?.value.trim() || '',
    NINEROUTER_FETCH_MODEL: document.getElementById('cfg-fetch-model')?.value.trim() || '',
    NINEROUTER_IMAGE_MODEL: document.getElementById('cfg-image-model')?.value.trim() || '',
    LLM_TIMEOUT: isNaN(timeoutVal) ? 300 : timeoutVal,
    VERIFICATION_ROUNDS: isNaN(vRoundsVal) ? 3 : vRoundsVal,
    PURE_PIL_ACTIVE: true
  };

  const agentsPayload = {};
  AGENT_IDS.forEach(aid => {
    const mVal = document.getElementById(`cfg-agent-model-${aid}`)?.value.trim() || '';
    const tVal = document.getElementById(`cfg-agent-think-${aid}`)?.value || 'default';
    config[`AGENT_MODEL_${aid.toUpperCase()}`] = mVal;
    config[`AGENT_THINK_LEVEL_${aid.toUpperCase()}`] = tVal;
    agentsPayload[aid] = {
      model: mVal,
      think_level: tVal
    };
  });

  try {
    const res = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ config, agents: agentsPayload })
    });
    const data = await res.json();
    if (data.success) {
      const activeModel = data.model || chatModel;
      showToast(`Saved settings & applied models (Primary: ${activeModel})`, 'success');
      applyConfigAndMetadata(data);
      updateAllAgentStatusBadges(data);
      await loadConfigBadge();
    }
  } catch (err) {
    showToast(`Save failed: ${err.message}`, 'error');
  }
}
