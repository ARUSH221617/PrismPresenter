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
    }
  } catch (err) {
    console.error('Failed to load settings', err);
  }
}

async function saveConfigSettings() {
  const config = {
    NINEROUTER_URL: document.getElementById('cfg-url').value.trim(),
    NINEROUTER_KEY: document.getElementById('cfg-key').value.trim(),
    NINEROUTER_CHAT_MODEL: document.getElementById('cfg-chat-model').value.trim(),
    NINEROUTER_SEARCH_MODEL: document.getElementById('cfg-search-model').value.trim(),
    NINEROUTER_FETCH_MODEL: document.getElementById('cfg-fetch-model').value.trim(),
    NINEROUTER_IMAGE_MODEL: document.getElementById('cfg-image-model').value.trim(),
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
