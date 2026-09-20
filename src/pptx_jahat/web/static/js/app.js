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
      const file = files[0];
      if (file.name.endsWith('.docx')) {
        uploadDocxFile(file);
      } else {
        showToast('Please upload a Microsoft Word (.docx) document.', 'warning');
      }
    }
  }, false);
}

// -------------------------------------------------------------
// KEYBOARD HOTKEYS SYSTEM
// -------------------------------------------------------------
function setupKeyboardHotkeys() {
  document.addEventListener('keydown', (e) => {
    // Quick Save shortcut for Settings tab
    if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      if (activeTab === 'settings') {
        e.preventDefault();
        saveConfigSettings();
        return;
      }
    }

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
  } else if (tabId === 'manager') {
    loadManagerDecks();
  } else if (tabId === 'components') {
    loadComponentsCatalog();
  } else if (tabId === 'settings') {
    loadConfigSettings();
    if (isSettingsDirty) {
      const bar = document.getElementById('settings-floating-bar');
      if (bar) bar.classList.remove('hidden');
    }
  }

  if (tabId !== 'settings') {
    const bar = document.getElementById('settings-floating-bar');
    if (bar) bar.classList.add('hidden');
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
    loadConfigBadge();
  } catch (err) {
    console.error('Error loading initial data', err);
  }
}

async function loadConfigBadge() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    if (data.success && data.config) {
      const modelName = data.config.NINEROUTER_CHAT_MODEL ? data.config.NINEROUTER_CHAT_MODEL.split('/').pop() : 'Default';
      const badge = document.getElementById('model-badge');
      if (badge) {
        badge.innerText = `MODEL: ${modelName}`;
      }
      const geminiModel = document.getElementById('gemini-model-name');
      if (geminiModel) {
        geminiModel.innerText = modelName;
      }
    }
  } catch (e) {
    console.error(e);
  }
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
  uploadDocxFile(file);
}

async function uploadDocxFile(file) {
  const status = document.getElementById('upload-status');
  status.classList.remove('hidden');
  status.innerText = `Uploading ${file.name}...`;
  status.className = 'text-[11px] text-muted-foreground mt-1.5';

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/generator/upload', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      document.getElementById('gen-docx-path').value = data.file_path;
      document.getElementById('gen-output-path').value = data.suggested_output;
      status.innerHTML = `<span class="text-emerald-400 font-medium">✓ Uploaded: ${data.filename}</span>`;
      showToast(`Uploaded ${data.filename}`, 'success');
    } else {
      status.innerHTML = `<span class="text-destructive">Upload failed: ${data.error}</span>`;
      showToast(`Upload failed: ${data.error}`, 'error');
    }
  } catch (err) {
    status.innerHTML = `<span class="text-destructive">Upload error: ${err.message}</span>`;
    showToast(`Upload error: ${err.message}`, 'error');
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
  const docxPath = document.getElementById('gen-docx-path').value.trim();
  const templateName = document.getElementById('gen-template-select').value;
  const outputPath = document.getElementById('gen-output-path').value.trim();

  if (!docxPath) {
    showToast('Please select or upload a Word (.docx) document first.', 'warning');
    return;
  }

  const btnGen = document.getElementById('btn-generate-pptx');
  const btnOpen = document.getElementById('btn-open-ppt');
  const btnDownload = document.getElementById('btn-download-pptx');

  btnGen.disabled = true;
  btnOpen.disabled = true;
  btnDownload.disabled = true;

  setSystemStatus('SYNTHESIZING...', true);
  appendGenLog(`[*] Starting presentation generation for ${docxPath}`);

  try {
    const res = await fetch('/api/generator/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        docx_path: docxPath,
        template_name: templateName,
        output_path: outputPath
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
// 6. SETTINGS & CONFIGURATION (15 UI/UX IMPROVEMENTS)
// -------------------------------------------------------------
let originalConfigState = null;
let isSettingsDirty = false;

async function loadConfigSettings() {
  try {
    const res = await fetch('/api/config');
    const data = await res.json();
    if (data.success && data.config) {
      const cfg = data.config;

      const urlInput = document.getElementById('cfg-url');
      const keyInput = document.getElementById('cfg-key');
      const chatInput = document.getElementById('cfg-chat-model');
      const searchInput = document.getElementById('cfg-search-model');
      const fetchInput = document.getElementById('cfg-fetch-model');
      const imageInput = document.getElementById('cfg-image-model');
      const renderModeInput = document.getElementById('cfg-render-mode');
      const purePilInput = document.getElementById('cfg-pure-pil');

      if (urlInput) urlInput.value = cfg.NINEROUTER_URL || '';
      if (keyInput) keyInput.value = cfg.NINEROUTER_KEY || '';
      if (chatInput) chatInput.value = cfg.NINEROUTER_CHAT_MODEL || '';
      if (searchInput) searchInput.value = cfg.NINEROUTER_SEARCH_MODEL || '';
      if (fetchInput) fetchInput.value = cfg.NINEROUTER_FETCH_MODEL || '';
      if (imageInput) imageInput.value = cfg.NINEROUTER_IMAGE_MODEL || '';

      const renderMode = (cfg.RENDER_MODE || 'auto').toLowerCase();
      if (renderModeInput) renderModeInput.value = renderMode;
      selectRenderMode(renderMode, false);

      const purePilActive = cfg.PURE_PIL_ACTIVE !== undefined ? Boolean(cfg.PURE_PIL_ACTIVE) : true;
      if (purePilInput) purePilInput.checked = purePilActive;

      // Cache snapshot for dirty-state diffing
      originalConfigState = {
        NINEROUTER_URL: cfg.NINEROUTER_URL || '',
        NINEROUTER_KEY: cfg.NINEROUTER_KEY || '',
        NINEROUTER_CHAT_MODEL: cfg.NINEROUTER_CHAT_MODEL || '',
        NINEROUTER_SEARCH_MODEL: cfg.NINEROUTER_SEARCH_MODEL || '',
        NINEROUTER_FETCH_MODEL: cfg.NINEROUTER_FETCH_MODEL || '',
        NINEROUTER_IMAGE_MODEL: cfg.NINEROUTER_IMAGE_MODEL || '',
        RENDER_MODE: renderMode,
        PURE_PIL_ACTIVE: purePilActive
      };

      updateKeyStatusIndicator();
      updateCascadeWaterfall();
      hideFloatingDirtyBar();
      loadDiagnostics();
      refreshIcons();
    }
  } catch (err) {
    console.error('Failed to load settings', err);
    showToast(`Failed to load settings: ${err.message}`, 'error');
  }
}

function getFormConfigState() {
  const purePilInput = document.getElementById('cfg-pure-pil');
  return {
    NINEROUTER_URL: (document.getElementById('cfg-url')?.value || '').trim(),
    NINEROUTER_KEY: (document.getElementById('cfg-key')?.value || '').trim(),
    NINEROUTER_CHAT_MODEL: (document.getElementById('cfg-chat-model')?.value || '').trim(),
    NINEROUTER_SEARCH_MODEL: (document.getElementById('cfg-search-model')?.value || '').trim(),
    NINEROUTER_FETCH_MODEL: (document.getElementById('cfg-fetch-model')?.value || '').trim(),
    NINEROUTER_IMAGE_MODEL: (document.getElementById('cfg-image-model')?.value || '').trim(),
    RENDER_MODE: (document.getElementById('cfg-render-mode')?.value || 'auto').trim(),
    PURE_PIL_ACTIVE: purePilInput ? purePilInput.checked : true
  };
}

function markSettingsDirty() {
  if (!originalConfigState) return;
  const current = getFormConfigState();
  let diffCount = 0;

  for (const key of Object.keys(originalConfigState)) {
    if (current[key] !== originalConfigState[key]) {
      diffCount++;
    }
  }

  isSettingsDirty = diffCount > 0;
  const bar = document.getElementById('settings-floating-bar');
  const countSpan = document.getElementById('dirty-changes-count');

  if (isSettingsDirty && activeTab === 'settings') {
    if (bar) bar.classList.remove('hidden');
    if (countSpan) {
      countSpan.innerText = `${diffCount} unsaved configuration change${diffCount > 1 ? 's' : ''}`;
    }
  } else {
    hideFloatingDirtyBar();
  }
}

function hideFloatingDirtyBar() {
  isSettingsDirty = false;
  const bar = document.getElementById('settings-floating-bar');
  if (bar) bar.classList.add('hidden');
}

function revertConfigSettings() {
  if (!originalConfigState) return;

  const urlInput = document.getElementById('cfg-url');
  const keyInput = document.getElementById('cfg-key');
  const chatInput = document.getElementById('cfg-chat-model');
  const searchInput = document.getElementById('cfg-search-model');
  const fetchInput = document.getElementById('cfg-fetch-model');
  const imageInput = document.getElementById('cfg-image-model');
  const renderModeInput = document.getElementById('cfg-render-mode');
  const purePilInput = document.getElementById('cfg-pure-pil');

  if (urlInput) urlInput.value = originalConfigState.NINEROUTER_URL;
  if (keyInput) keyInput.value = originalConfigState.NINEROUTER_KEY;
  if (chatInput) chatInput.value = originalConfigState.NINEROUTER_CHAT_MODEL;
  if (searchInput) searchInput.value = originalConfigState.NINEROUTER_SEARCH_MODEL;
  if (fetchInput) fetchInput.value = originalConfigState.NINEROUTER_FETCH_MODEL;
  if (imageInput) imageInput.value = originalConfigState.NINEROUTER_IMAGE_MODEL;
  if (renderModeInput) renderModeInput.value = originalConfigState.RENDER_MODE;
  selectRenderMode(originalConfigState.RENDER_MODE, false);
  if (purePilInput) purePilInput.checked = originalConfigState.PURE_PIL_ACTIVE;

  updateKeyStatusIndicator();
  updateCascadeWaterfall();
  hideFloatingDirtyBar();
  showToast('Configuration changes reverted', 'info');
}

async function saveConfigSettings() {
  const current = getFormConfigState();

  // Loading UI feedback
  const saveBtn = document.getElementById('btn-save-settings');
  const saveIcon = document.getElementById('save-btn-icon');
  const saveSpinner = document.getElementById('save-btn-spinner');
  const saveText = document.getElementById('save-btn-text');

  if (saveBtn) saveBtn.disabled = true;
  if (saveIcon) saveIcon.classList.add('hidden');
  if (saveSpinner) saveSpinner.classList.remove('hidden');
  if (saveText) saveText.innerText = 'Saving...';

  try {
    const res = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ config: current })
    });
    const data = await res.json();
    if (data.success) {
      originalConfigState = { ...current };
      hideFloatingDirtyBar();
      showToast('Configuration saved to .env (backup created)', 'success');
      loadConfigBadge();
      loadDiagnostics();
    } else {
      showToast(`Save failed: ${data.error || 'Unknown error'}`, 'error');
    }
  } catch (err) {
    showToast(`Save failed: ${err.message}`, 'error');
  } finally {
    if (saveBtn) saveBtn.disabled = false;
    if (saveIcon) saveIcon.classList.remove('hidden');
    if (saveSpinner) saveSpinner.classList.add('hidden');
    if (saveText) saveText.innerText = 'Save Settings (.env)';
    refreshIcons();
  }
}

// -------------------------------------------------------------
// GATEWAY CONNECTIVITY & API KEY MANAGEMENT
// -------------------------------------------------------------
async function testGatewayConnection() {
  const urlInput = document.getElementById('cfg-url');
  const keyInput = document.getElementById('cfg-key');
  const pingIcon = document.getElementById('ping-icon');
  const pingSpinner = document.getElementById('ping-spinner');
  const pingBtn = document.getElementById('btn-test-connection');
  const resultBox = document.getElementById('cfg-ping-result');
  const resultText = document.getElementById('cfg-ping-text');
  const latencyText = document.getElementById('cfg-ping-latency');
  const statusPill = document.getElementById('gateway-status-pill');

  let targetUrl = (urlInput ? urlInput.value.trim() : '');
  if (!targetUrl) targetUrl = 'http://localhost:20128';
  const apiKey = (keyInput ? keyInput.value.trim() : '');

  if (pingBtn) pingBtn.disabled = true;
  if (pingIcon) pingIcon.classList.add('hidden');
  if (pingSpinner) pingSpinner.classList.remove('hidden');

  try {
    const res = await fetch('/api/config/ping', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url: targetUrl, key: apiKey })
    });
    const data = await res.json();

    if (resultBox) resultBox.classList.remove('hidden');

    if (data.success) {
      const modelInfo = (data.model_count !== null && data.model_count !== undefined) ? ` • ${data.model_count} models discovered` : '';
      if (resultBox) {
        resultBox.className = 'mt-2.5 p-2.5 rounded-md border border-emerald-500/40 bg-emerald-500/10 text-xs flex items-center justify-between';
      }
      if (resultText) {
        resultText.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-500 shrink-0"></span><span class="text-emerald-400 font-medium">Gateway Online (${data.message})${modelInfo}</span>`;
      }
      if (latencyText) {
        latencyText.innerText = `${data.latency_ms}ms latency`;
      }
      if (statusPill) {
        statusPill.className = 'shadcn-badge shadcn-badge-outline text-emerald-400 border-emerald-500/30 gap-1.5 font-mono text-[10px] py-0.5';
        statusPill.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span> Online (${data.latency_ms}ms)`;
      }
      showToast(`Gateway connection verified (${data.latency_ms}ms)`, 'success');
    } else {
      if (resultBox) {
        resultBox.className = 'mt-2.5 p-2.5 rounded-md border border-destructive/40 bg-destructive/10 text-xs flex items-center justify-between';
      }
      if (resultText) {
        resultText.innerHTML = `<span class="w-2 h-2 rounded-full bg-destructive shrink-0"></span><span class="text-destructive font-medium">${escapeHtml(data.error || 'Connection failed')}</span>`;
      }
      if (latencyText) {
        latencyText.innerText = `${data.latency_ms || 0}ms`;
      }
      if (statusPill) {
        statusPill.className = 'shadcn-badge shadcn-badge-outline text-destructive border-destructive/30 gap-1.5 font-mono text-[10px] py-0.5';
        statusPill.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-destructive"></span> Offline`;
      }
      showToast(`Gateway connection failed: ${data.error || 'Unreachable'}`, 'error');
    }
  } catch (err) {
    if (resultBox) {
      resultBox.classList.remove('hidden');
      resultBox.className = 'mt-2.5 p-2.5 rounded-md border border-destructive/40 bg-destructive/10 text-xs flex items-center justify-between';
    }
    if (resultText) {
      resultText.innerHTML = `<span class="w-2 h-2 rounded-full bg-destructive shrink-0"></span><span class="text-destructive font-medium">Network error: ${escapeHtml(err.message)}</span>`;
    }
    if (statusPill) {
      statusPill.className = 'shadcn-badge shadcn-badge-outline text-destructive border-destructive/30 gap-1.5 font-mono text-[10px] py-0.5';
      statusPill.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-destructive"></span> Error`;
    }
  } finally {
    if (pingBtn) pingBtn.disabled = false;
    if (pingIcon) pingIcon.classList.remove('hidden');
    if (pingSpinner) pingSpinner.classList.add('hidden');
    refreshIcons();
  }
}

function toggleKeyVisibility() {
  const keyInput = document.getElementById('cfg-key');
  const eyeIcon = document.getElementById('key-icon-eye');
  const eyeOffIcon = document.getElementById('key-icon-eye-off');
  if (!keyInput) return;

  if (keyInput.type === 'password') {
    keyInput.type = 'text';
    if (eyeIcon) eyeIcon.classList.add('hidden');
    if (eyeOffIcon) eyeOffIcon.classList.remove('hidden');
  } else {
    keyInput.type = 'password';
    if (eyeIcon) eyeIcon.classList.remove('hidden');
    if (eyeOffIcon) eyeOffIcon.classList.add('hidden');
  }
}

async function pasteKeyFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    if (text) {
      const keyInput = document.getElementById('cfg-key');
      if (keyInput) {
        keyInput.value = text.trim();
        updateKeyStatusIndicator();
        markSettingsDirty();
        showToast('Pasted API key from clipboard', 'info');
      }
    }
  } catch (e) {
    showToast('Clipboard access was blocked or empty', 'warning');
  }
}

function updateKeyStatusIndicator() {
  const keyInput = document.getElementById('cfg-key');
  const statusSpan = document.getElementById('cfg-key-status');
  if (!statusSpan) return;

  const val = (keyInput?.value || '').trim();
  if (val.length > 0) {
    const preview = val.length > 6 ? `(starts with ${escapeHtml(val.substring(0, 4))}••••)` : '(active)';
    statusSpan.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-emerald-400"></span> <span class="text-foreground">Key configured ${preview}</span>`;
  } else {
    statusSpan.innerHTML = `<span class="w-1.5 h-1.5 rounded-full bg-muted-foreground/60"></span> <span>No key configured (local gateway mode)</span>`;
  }
}

function sanitizeGatewayUrl() {
  const urlInput = document.getElementById('cfg-url');
  if (!urlInput) return;
  let val = urlInput.value.trim();
  if (val.length > 0) {
    if (!val.startsWith('http://') && !val.startsWith('https://')) {
      val = 'http://' + val;
    }
    val = val.replace(/\/+$/, '');
    urlInput.value = val;
    markSettingsDirty();
  }
}

// -------------------------------------------------------------
// PRESETS & MODEL CHIPS
// -------------------------------------------------------------
function applyPreset(presetType) {
  const presets = {
    recommended: {
      url: 'http://localhost:20128',
      chat: 'ag/gemini-3.7-flash-high',
      search: 'tavily',
      fetch: 'jina-reader',
      image: 'gemini/gemini-3-pro-image-preview',
      mode: 'auto',
      pil: true
    },
    fast: {
      url: 'http://localhost:20128',
      chat: 'ag/gemini-3.7-flash-high',
      search: 'brave-search',
      fetch: 'jina-reader',
      image: 'gemini/gemini-3-pro-image-preview',
      mode: 'web',
      pil: true
    },
    reasoning: {
      url: 'http://localhost:20128',
      chat: 'aval/claude-sonnet-4-6',
      search: 'exa_ai-search',
      fetch: 'firecrawl-search',
      image: 'gemini/gemini-3-pro-image-preview',
      mode: 'native',
      pil: false
    },
    local: {
      url: 'http://localhost:20128',
      key: '',
      chat: 'ag/gemini-3.7-flash-high',
      search: 'tavily',
      fetch: 'jina-reader',
      image: 'gemini/gemini-3-pro-image-preview',
      mode: 'auto',
      pil: true
    }
  };

  const p = presets[presetType];
  if (!p) return;

  const urlInput = document.getElementById('cfg-url');
  const chatInput = document.getElementById('cfg-chat-model');
  const searchInput = document.getElementById('cfg-search-model');
  const fetchInput = document.getElementById('cfg-fetch-model');
  const imageInput = document.getElementById('cfg-image-model');
  const purePilInput = document.getElementById('cfg-pure-pil');

  if (urlInput) urlInput.value = p.url;
  if (p.key !== undefined) {
    const keyInput = document.getElementById('cfg-key');
    if (keyInput) keyInput.value = p.key;
  }
  if (chatInput) chatInput.value = p.chat;
  if (searchInput) searchInput.value = p.search;
  if (fetchInput) fetchInput.value = p.fetch;
  if (imageInput) imageInput.value = p.image;
  if (purePilInput) purePilInput.checked = p.pil;

  selectRenderMode(p.mode);
  updateCascadeWaterfall();
  updateKeyStatusIndicator();
  markSettingsDirty();

  showToast(`Applied ${presetType.toUpperCase()} preset profile`, 'info');
}

function setFieldChip(fieldId, value) {
  const el = document.getElementById(fieldId);
  if (!el) return;
  el.value = value;
  markSettingsDirty();

  el.classList.add('ring-2', 'ring-primary');
  setTimeout(() => el.classList.remove('ring-2', 'ring-primary'), 350);
}

// -------------------------------------------------------------
// RENDER CASCADE SELECTOR & VISUAL WATERFALL
// -------------------------------------------------------------
function selectRenderMode(mode, triggerDirty = true) {
  const hiddenInput = document.getElementById('cfg-render-mode');
  if (hiddenInput) hiddenInput.value = mode;

  document.querySelectorAll('.render-mode-card').forEach(card => {
    card.classList.remove('selected');
  });

  const targetCard = document.getElementById(`render-mode-${mode}`);
  if (targetCard) targetCard.classList.add('selected');

  updateCascadeWaterfall();
  if (triggerDirty) markSettingsDirty();
}

function updateCascadeWaterfall() {
  const purePilInput = document.getElementById('cfg-pure-pil');
  const modeInput = document.getElementById('cfg-render-mode');
  const flowCom = document.getElementById('flow-step-com');
  const flowWeb = document.getElementById('flow-step-web');
  const flowPil = document.getElementById('flow-step-pil');
  const pilStatus = document.getElementById('flow-pil-status');
  const warningText = document.getElementById('flow-warning-text');

  const mode = (modeInput?.value || 'auto').toLowerCase();
  const pilActive = purePilInput ? purePilInput.checked : true;

  if (flowCom && flowWeb && flowPil) {
    flowCom.className = 'cascade-flow-step active';
    flowWeb.className = 'cascade-flow-step active';
    flowPil.className = 'cascade-flow-step active';

    if (mode === 'native') {
      flowWeb.classList.add('disabled-step');
      flowPil.classList.add('disabled-step');
      if (pilStatus) pilStatus.innerText = 'Bypassed (COM only)';
      if (warningText) warningText.innerText = 'Strict Native COM mode: requires Microsoft PowerPoint installed on host.';
    } else if (mode === 'web') {
      flowCom.classList.add('disabled-step');
      flowPil.classList.add('disabled-step');
      if (pilStatus) pilStatus.innerText = 'Bypassed (Web only)';
      if (warningText) warningText.innerText = 'Direct Web Vector Engine mode active.';
    } else if (mode === 'pil') {
      flowCom.classList.add('disabled-step');
      flowWeb.classList.add('disabled-step');
      if (pilStatus) pilStatus.innerText = 'Active Standalone';
      if (warningText) warningText.innerText = 'Pure PIL Fallback mode: generates slide preview raster images directly.';
    } else {
      if (!pilActive) {
        flowPil.classList.add('disabled-step');
        if (pilStatus) {
          pilStatus.innerText = 'Disabled (Strict)';
          pilStatus.className = 'text-[10px] text-destructive font-mono';
        }
        if (warningText) warningText.innerText = 'Warning: Pure PIL fallback is disabled. If PowerPoint COM & Web render fails, generation will halt with an error.';
      } else {
        if (pilStatus) {
          pilStatus.innerText = 'Active (OK)';
          pilStatus.className = 'text-[10px] text-emerald-400 font-mono';
        }
        if (warningText) warningText.innerText = '';
      }
    }
  }
}

// -------------------------------------------------------------
// DIAGNOSTICS & CACHE CLEANER
// -------------------------------------------------------------
async function loadDiagnostics() {
  const diagOs = document.getElementById('diag-os');
  const diagCom = document.getElementById('diag-com');
  const diagStorage = document.getElementById('diag-storage');
  const diagComponents = document.getElementById('diag-components');
  const diagTemplates = document.getElementById('diag-templates');
  const refreshIcon = document.getElementById('diag-refresh-icon');

  if (refreshIcon) refreshIcon.classList.add('animate-spin');

  try {
    const res = await fetch('/api/config/diagnostics');
    const data = await res.json();
    if (data.success) {
      if (diagOs) {
        diagOs.innerText = `${data.platform.system} (${data.platform.os}) • Py ${data.platform.python}`;
      }
      if (diagCom) {
        if (data.com_engine.status === 'available') {
          diagCom.innerHTML = `<span class="shadcn-badge shadcn-badge-outline text-emerald-400 border-emerald-500/30 text-[10px] py-0 px-1.5"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500 mr-1 inline-block"></span> COM Ready</span>`;
        } else {
          diagCom.innerHTML = `<span class="shadcn-badge shadcn-badge-outline text-muted-foreground text-[10px] py-0 px-1.5">Unavailable</span>`;
        }
      }
      if (diagStorage) {
        diagStorage.innerText = `${data.storage.output_files_count} files • ${data.storage.output_size_mb} MB`;
      }
      if (diagComponents) {
        diagComponents.innerText = `${data.storage.components_count} Archetypes`;
      }
      if (diagTemplates) {
        diagTemplates.innerText = `${data.storage.templates_count} PPTX Decks`;
      }
    }
  } catch (err) {
    console.warn('Diagnostics fetch failed', err);
  } finally {
    if (refreshIcon) refreshIcon.classList.remove('animate-spin');
  }
}

async function cleanRenderCache() {
  const cleanBtn = document.getElementById('btn-clean-cache');
  const cleanIcon = document.getElementById('clean-cache-icon');
  const cleanSpinner = document.getElementById('clean-cache-spinner');

  if (cleanBtn) cleanBtn.disabled = true;
  if (cleanIcon) cleanIcon.classList.add('hidden');
  if (cleanSpinner) cleanSpinner.classList.remove('hidden');

  try {
    const res = await fetch('/api/config/clean-cache', { method: 'POST' });
    const data = await res.json();
    if (data.success) {
      showToast(`Cache cleaned: freed ${data.freed_mb} MB (${data.cleaned_count} files removed)`, 'success');
      loadDiagnostics();
    } else {
      showToast(`Cache cleaning error: ${data.error || 'Failed'}`, 'error');
    }
  } catch (err) {
    showToast(`Failed to clean cache: ${err.message}`, 'error');
  } finally {
    if (cleanBtn) cleanBtn.disabled = false;
    if (cleanIcon) cleanIcon.classList.remove('hidden');
    if (cleanSpinner) cleanSpinner.classList.add('hidden');
    refreshIcons();
  }
}

// -------------------------------------------------------------
// DUAL MODE: FORM VS RAW .ENV SYNTAX EDITOR
// -------------------------------------------------------------
async function switchSettingsView(mode) {
  const formView = document.getElementById('settings-form-view');
  const rawView = document.getElementById('settings-raw-view');
  const formBtn = document.getElementById('btn-view-form');
  const rawBtn = document.getElementById('btn-view-raw');

  if (mode === 'raw') {
    if (formView) formView.classList.add('hidden');
    if (rawView) rawView.classList.remove('hidden');
    if (formBtn) formBtn.classList.remove('active');
    if (rawBtn) rawBtn.classList.add('active');

    try {
      const res = await fetch('/api/config/raw');
      const data = await res.json();
      const rawText = document.getElementById('cfg-raw-text');
      if (rawText) rawText.value = data.raw || '';
    } catch (e) {
      console.warn('Failed to load raw .env', e);
    }
  } else {
    if (rawView) rawView.classList.add('hidden');
    if (formView) formView.classList.remove('hidden');
    if (rawBtn) rawBtn.classList.remove('active');
    if (formBtn) formBtn.classList.add('active');
  }
  refreshIcons();
}

function copyRawConfig() {
  const rawText = document.getElementById('cfg-raw-text');
  if (!rawText) return;
  navigator.clipboard.writeText(rawText.value).then(() => {
    showToast('Copied .env configuration to clipboard', 'success');
  }).catch(() => {
    showToast('Failed to copy to clipboard', 'warning');
  });
}

function syncRawToForm() {
  const rawText = document.getElementById('cfg-raw-text');
  if (!rawText) return;

  const lines = rawText.value.split('\n');
  const parsed = {};
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx !== -1) {
      const k = trimmed.substring(0, eqIdx).trim();
      const v = trimmed.substring(eqIdx + 1).trim();
      parsed[k] = v;
    }
  }

  if (parsed.NINEROUTER_URL !== undefined) {
    const el = document.getElementById('cfg-url');
    if (el) el.value = parsed.NINEROUTER_URL;
  }
  if (parsed.NINEROUTER_KEY !== undefined) {
    const el = document.getElementById('cfg-key');
    if (el) el.value = parsed.NINEROUTER_KEY;
  }
  if (parsed.NINEROUTER_CHAT_MODEL !== undefined) {
    const el = document.getElementById('cfg-chat-model');
    if (el) el.value = parsed.NINEROUTER_CHAT_MODEL;
  }
  if (parsed.NINEROUTER_SEARCH_MODEL !== undefined) {
    const el = document.getElementById('cfg-search-model');
    if (el) el.value = parsed.NINEROUTER_SEARCH_MODEL;
  }
  if (parsed.NINEROUTER_FETCH_MODEL !== undefined) {
    const el = document.getElementById('cfg-fetch-model');
    if (el) el.value = parsed.NINEROUTER_FETCH_MODEL;
  }
  if (parsed.NINEROUTER_IMAGE_MODEL !== undefined) {
    const el = document.getElementById('cfg-image-model');
    if (el) el.value = parsed.NINEROUTER_IMAGE_MODEL;
  }
  if (parsed.RENDER_MODE !== undefined) {
    selectRenderMode(parsed.RENDER_MODE.toLowerCase());
  }
  if (parsed.PURE_PIL_ACTIVE !== undefined) {
    const el = document.getElementById('cfg-pure-pil');
    if (el) el.checked = ['1', 'true', 'yes', 'on'].includes(parsed.PURE_PIL_ACTIVE.toLowerCase());
  }

  updateCascadeWaterfall();
  updateKeyStatusIndicator();
  markSettingsDirty();
  switchSettingsView('form');
  showToast('Synced variables into form fields', 'info');
}

async function saveRawConfig() {
  const rawText = document.getElementById('cfg-raw-text');
  if (!rawText) return;

  try {
    const res = await fetch('/api/config/raw', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ raw: rawText.value })
    });
    const data = await res.json();
    if (data.success) {
      showToast('Raw .env saved and reloaded', 'success');
      loadConfigSettings();
      loadConfigBadge();
    } else {
      showToast('Save failed', 'error');
    }
  } catch (err) {
    showToast(`Save failed: ${err.message}`, 'error');
  }
}

// -------------------------------------------------------------
// RESET DEFAULTS MODAL
// -------------------------------------------------------------
function openResetDefaultsModal() {
  const modal = document.getElementById('reset-defaults-modal');
  if (modal) modal.classList.remove('hidden');
  refreshIcons();
}

function closeResetDefaultsModal() {
  const modal = document.getElementById('reset-defaults-modal');
  if (modal) modal.classList.add('hidden');
}

function confirmResetDefaults() {
  closeResetDefaultsModal();
  applyPreset('recommended');
  showToast('Recommended defaults applied. Click Save to persist.', 'info');
}
