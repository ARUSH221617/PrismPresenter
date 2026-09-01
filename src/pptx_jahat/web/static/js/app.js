// PPTX Jahat Web SPA Client Application Logic

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

// DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    lucide.createIcons();
  }
  loadInitialData();
});

function refreshIcons() {
  if (window.lucide) {
    setTimeout(() => lucide.createIcons(), 50);
  }
}

// -------------------------------------------------------------
// TAB SWITCHING
// -------------------------------------------------------------
function switchTab(tabId) {
  activeTab = tabId;
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active', 'border-red-600', 'text-white');
    btn.classList.add('border-transparent', 'text-gray-400');
  });

  const activeBtn = document.getElementById(`tab-btn-${tabId}`);
  if (activeBtn) {
    activeBtn.classList.add('active', 'border-red-600', 'text-white');
    activeBtn.classList.remove('border-transparent', 'text-gray-400');
  }

  document.querySelectorAll('.tab-content').forEach(content => {
    content.classList.add('hidden');
  });

  const activeContent = document.getElementById(`tab-content-${tabId}`);
  if (activeContent) {
    activeContent.classList.remove('hidden');
  }

  // Trigger tab-specific loaders
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
  document.querySelectorAll('.gen-subtab-btn').forEach(btn => {
    btn.classList.remove('active', 'bg-red-600', 'text-white');
    btn.classList.add('text-gray-400');
  });

  const activeBtn = document.getElementById(`subtab-btn-${subTabId}`);
  if (activeBtn) {
    activeBtn.classList.add('active', 'bg-red-600', 'text-white');
    activeBtn.classList.remove('text-gray-400');
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
    badge.className = 'px-2.5 py-1 rounded-full text-xs font-semibold bg-red-600 text-white flex items-center gap-1.5 shadow-sm';
  } else {
    badge.className = 'px-2.5 py-1 rounded-full text-xs font-semibold bg-red-950/60 border border-red-800/70 text-red-300 flex items-center gap-1.5 shadow-sm';
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
      const modelName = data.config.NINEROUTER_CHAT_MODEL.split('/').pop();
      document.getElementById('model-badge').innerText = `MODEL: ${modelName}`;
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
    select.innerHTML = '<option value="">All Templates (Global AI Matching)</option>';

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

  const status = document.getElementById('upload-status');
  status.classList.remove('hidden');
  status.innerText = `Uploading ${file.name}...`;

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
      status.innerText = `Uploaded: ${data.filename}`;
      status.className = 'text-[11px] text-emerald-400 mt-1';
    } else {
      status.innerText = `Upload failed: ${data.error}`;
      status.className = 'text-[11px] text-red-400 mt-1';
    }
  } catch (err) {
    status.innerText = `Upload error: ${err.message}`;
    status.className = 'text-[11px] text-red-400 mt-1';
  }
}

function appendGenLog(msg, type = 'info') {
  const logs = document.getElementById('gen-console-logs');
  const line = document.createElement('div');

  if (msg.startsWith('[*]')) {
    line.className = 'text-sky-400 font-semibold';
  } else if (msg.startsWith('[✓]') || msg.includes('SUCCESS')) {
    line.className = 'text-emerald-400 font-semibold';
  } else if (msg.startsWith('[!]') || msg.includes('ERROR')) {
    line.className = 'text-red-400 font-bold';
  } else {
    line.className = 'text-gray-300';
  }

  line.innerText = msg;
  logs.appendChild(line);
  logs.scrollTop = logs.scrollHeight;
}

function clearGenLog() {
  document.getElementById('gen-console-logs').innerHTML = '';
}

async function startPresentationGeneration() {
  const docxPath = document.getElementById('gen-docx-path').value.trim();
  const templateName = document.getElementById('gen-template-select').value;
  const outputPath = document.getElementById('gen-output-path').value.trim();

  if (!docxPath) {
    alert('Please select or upload a Word (.docx) document first.');
    return;
  }

  const btnGen = document.getElementById('btn-generate-pptx');
  const btnOpen = document.getElementById('btn-open-ppt');
  const btnDownload = document.getElementById('btn-download-pptx');

  btnGen.disabled = true;
  btnOpen.disabled = true;
  btnDownload.disabled = true;

  setSystemStatus('● GENERATING PRESENTATION...', true);
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
    setSystemStatus('● GENERATION FAILED');
    btnGen.disabled = false;
    alert(`Generation Error: ${err.message}`);
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

    btnOpen.disabled = false;
    btnDownload.disabled = false;
    btnGen.disabled = false;

    setSystemStatus('● BUILD READY');
    appendGenLog(`[✓] Completed: ${d.filename}`, 'success');
  });

  evtSource.addEventListener('error', (e) => {
    try {
      const d = JSON.parse(e.data);
      appendGenLog(`[!] Error: ${d.error}`);
    } catch (_) {}
    setSystemStatus('● BUILD FAILED');
    btnGen.disabled = false;
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
      badge.innerText = '⚡ Native PowerPoint';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 border border-emerald-700 text-emerald-300';
    } else if (engineName.includes('Web')) {
      badge.innerText = '🌐 Web Render Engine';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-sky-950 border border-sky-700 text-sky-300';
    } else {
      badge.innerText = '🎨 Pure PIL';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-red-950 border border-red-700 text-red-300';
    }
  }
}

function navGenSlide(dir) {
  if (genSlideIdx + dir >= 0 && genSlideIdx + dir < genSlides.length) {
    genSlideIdx += dir;
    updateGenSlideDisplay();
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
    counter.innerText = 'No visual screenshot loaded';
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  ph.classList.add('hidden');
  img.classList.remove('hidden');
  img.src = visualSlides[visualSlideIdx].data_url;
  counter.innerText = `Visual Screenshot ${visualSlideIdx + 1} of ${visualSlides.length}`;
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
  counter.innerText = `AI Payload ${aiTestIdx + 1}/${aiTestImages.length} • ${cur.template_file} [Slide ${cur.slide_index + 1}] (${cur.archetype})`;
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
  } catch (err) {
    alert(`Could not launch presentation: ${err.message}`);
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
      tbody.innerHTML = '<tr><td colspan="4" class="py-4 text-center text-gray-500 italic">No templates found in data/.</td></tr>';
      return;
    }

    templatesList.forEach((tpl, idx) => {
      const tr = document.createElement('tr');
      tr.className = `cursor-pointer hover:bg-[#1b1e2b] transition ${selectedTemplateName === tpl.filename ? 'bg-[#1e2230]' : ''}`;
      tr.onclick = () => selectTemplateItem(tpl.filename);

      const statusTag = tpl.is_analyzed
        ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-950 text-emerald-400 border border-emerald-800">✓ Analyzed</span>'
        : '<span class="px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-800 text-gray-400 border border-gray-700">○ Pending</span>';

      tr.innerHTML = `
        <td class="py-2.5 px-3 font-medium text-white font-mono">${tpl.filename}</td>
        <td class="py-2.5 px-3 text-center text-gray-300 font-mono">${tpl.slide_count}</td>
        <td class="py-2.5 px-3 text-center">${statusTag}</td>
        <td class="py-2.5 px-3 text-gray-300 truncate max-w-[140px]">${tpl.style} • ${tpl.purpose}</td>
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
      <div class="font-bold text-white text-xs">${tpl.filename}</div>
      <div class="text-gray-400 text-[11px]">Slides: ${tpl.slide_count} | Dimensions: ${tpl.dimensions}</div>
      <div class="pt-2 border-t border-[#272b3c] space-y-1">
        <div><span class="text-red-400 font-semibold">🎯 Purpose:</span> ${tpl.purpose || 'Not analyzed'}</div>
        <div><span class="text-sky-400 font-semibold">🎨 Style:</span> ${tpl.style || 'Not analyzed'}</div>
        <div><span class="text-yellow-400 font-semibold">📝 Brief:</span> ${tpl.brief || 'Click Analyze to generate'}</div>
      </div>
    `;

    // Load previews for selected template
    loadTemplateSlidePreviews(tpl.file_path);
  }
}

function loadTemplatesListVisuals() {
  document.querySelectorAll('#templates-table-body tr').forEach(tr => {
    if (tr.innerText.includes(selectedTemplateName)) {
      tr.classList.add('bg-[#1e2230]');
    } else {
      tr.classList.remove('bg-[#1e2230]');
    }
  });
}

async function loadTemplateSlidePreviews(filePath) {
  const ph = document.getElementById('tpl-slide-placeholder');
  const img = document.getElementById('tpl-slide-img');
  const counter = document.getElementById('tpl-slide-counter');

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
      alert('Template Intelligence Notes (data/NOTE.md) saved successfully!');
      loadTemplatesList();
    }
  } catch (err) {
    alert(`Save error: ${err.message}`);
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
    alert('Please select a template from the table first.');
    return;
  }

  const btnSel = document.getElementById('btn-analyze-sel');
  const btnAll = document.getElementById('btn-analyze-all');
  btnSel.disabled = true;
  btnAll.disabled = true;

  appendAnalyzeLog(`\n[*] Starting AI analysis for template: ${selectedTemplateName}`);

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
    evtSource.addEventListener('completed', (e) => {
      appendAnalyzeLog(`[✓] Finished analysis for ${selectedTemplateName}`);
      loadTemplatesList();
      loadNoteMd();
      btnSel.disabled = false;
      btnAll.disabled = false;
    });
    evtSource.addEventListener('error', (e) => {
      appendAnalyzeLog(`[!] Error analyzing template`);
      btnSel.disabled = false;
      btnAll.disabled = false;
    });
    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAnalyzeLog(`[!] Analysis failed: ${err.message}`);
    btnSel.disabled = false;
    btnAll.disabled = false;
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

    evtSource.addEventListener('completed', (e) => {
      appendAnalyzeLog(`[✓] Completed batch analysis of all templates.`);
      progBar.style.width = '100%';
      progPct.innerText = '100%';
      loadTemplatesList();
      loadNoteMd();
      btnSel.disabled = false;
      btnAll.disabled = false;
    });

    evtSource.addEventListener('error', () => {
      appendAnalyzeLog(`[!] Error in batch template analysis.`);
      btnSel.disabled = false;
      btnAll.disabled = false;
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAnalyzeLog(`[!] Batch analysis failed: ${err.message}`);
    btnSel.disabled = false;
    btnAll.disabled = false;
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
      genBody.innerHTML = '<tr><td colspan="3" class="py-3 text-center text-gray-500 italic">No generated decks found.</td></tr>';
    } else {
      mgrGeneratedList.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = `cursor-pointer hover:bg-[#1b1e2b] transition ${mgrSelectedFile?.file_path === item.file_path ? 'bg-[#1e2230]' : ''}`;
        tr.onclick = () => selectManagerFile(item);
        tr.innerHTML = `
          <td class="py-2 px-3 font-medium text-white truncate max-w-[200px] font-mono">${item.filename}</td>
          <td class="py-2 px-3 text-center text-gray-300 font-mono text-[11px]">${item.size}</td>
          <td class="py-2 px-3 text-center text-gray-400 font-mono text-[11px]">${item.modified}</td>
        `;
        genBody.appendChild(tr);
      });
    }

    // Render Reference table
    const refBody = document.getElementById('mgr-ref-table-body');
    refBody.innerHTML = '';
    if (mgrReferenceList.length === 0) {
      refBody.innerHTML = '<tr><td colspan="3" class="py-3 text-center text-gray-500 italic">No reference templates found.</td></tr>';
    } else {
      mgrReferenceList.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = `cursor-pointer hover:bg-[#1b1e2b] transition ${mgrSelectedFile?.file_path === item.file_path ? 'bg-[#1e2230]' : ''}`;
        tr.onclick = () => selectManagerFile(item);
        tr.innerHTML = `
          <td class="py-2 px-3 font-medium text-white truncate max-w-[200px] font-mono">${item.filename}</td>
          <td class="py-2 px-3 text-center text-gray-300 font-mono text-[11px]">${item.size}</td>
          <td class="py-2 px-3 text-center text-gray-400 font-mono text-[11px]">${item.modified}</td>
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
      alert(`Imported template: ${data.filename}`);
      loadManagerDecks();
    } else {
      alert(`Import failed: ${data.error}`);
    }
  } catch (err) {
    alert(`Import error: ${err.message}`);
  }
}

function selectManagerFile(item) {
  mgrSelectedFile = item;
  document.getElementById('mgr-sel-name').innerText = `📄 ${item.filename}`;
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
      badge.innerText = '⚡ Native PowerPoint';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-emerald-950 border border-emerald-700 text-emerald-300';
    } else if (engineName.includes('Web')) {
      badge.innerText = '🌐 Web Render Engine';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-sky-950 border border-sky-700 text-sky-300';
    } else {
      badge.innerText = '🎨 Pure PIL';
      badge.className = 'text-[10px] font-mono px-2 py-0.5 rounded bg-red-950 border border-red-700 text-red-300';
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
  } catch (err) {
    alert(`Could not open file: ${err.message}`);
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
      alert(`Integrity Check Passed!\nAuto-healing verified: ${data.filename}`);
      loadManagerDecks();
    }
  } catch (err) {
    alert(`Verify error: ${err.message}`);
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
      alert(`Created duplicate: ${data.filename}`);
      loadManagerDecks();
    }
  } catch (err) {
    alert(`Duplicate error: ${err.message}`);
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
      alert(`Renamed to: ${data.filename}`);
      loadManagerDecks();
    } else {
      alert(`Rename failed: ${data.error}`);
    }
  } catch (err) {
    alert(`Rename error: ${err.message}`);
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
      alert('Presentation deleted.');
      mgrSelectedFile = null;
      loadManagerDecks();
    }
  } catch (err) {
    alert(`Delete error: ${err.message}`);
  }
}

// -------------------------------------------------------------
// 4. COMPONENT CATALOG
// -------------------------------------------------------------
async function loadComponentsCatalog() {
  try {
    const res = await fetch('/api/components/catalog');
    const data = await res.json();
    document.getElementById('comp-count-badge').innerText = `Extracted Components: ${data.count}`;
    document.getElementById('components-json-viewer').value = JSON.stringify(data.catalog, null, 2);
  } catch (err) {
    console.error('Error loading component catalog', err);
  }
}

async function runComponentsExtraction() {
  const viewer = document.getElementById('components-json-viewer');
  viewer.value = '[*] Scanning data/*.pptx templates and extracting components...';

  try {
    const res = await fetch('/api/components/extract', { method: 'POST' });
    const data = await res.json();
    if (data.success) {
      document.getElementById('comp-count-badge').innerText = `Extracted Components: ${data.count}`;
      viewer.value = JSON.stringify(data.catalog, null, 2);
      alert(`Component extraction complete!\nDiscovered ${data.count} visual components.`);
    }
  } catch (err) {
    alert(`Extraction error: ${err.message}`);
  }
}

// -------------------------------------------------------------
// 5. AUTONOMOUS AI AGENT
// -------------------------------------------------------------
function appendAgentLog(msg, isPrompt = false) {
  const logs = document.getElementById('agent-console-logs');
  const line = document.createElement('div');

  if (isPrompt) {
    line.className = 'text-red-400 font-bold border-l-2 border-red-500 pl-2 my-2';
  } else if (msg.includes('Final Answer') || msg.includes('SUCCESS')) {
    line.className = 'text-emerald-400 font-semibold';
  } else {
    line.className = 'text-gray-300';
  }

  line.innerText = msg;
  logs.appendChild(line);
  logs.scrollTop = logs.scrollHeight;
}

function clearAgentLogs() {
  document.getElementById('agent-console-logs').innerHTML = '';
}

async function sendAgentPrompt() {
  const input = document.getElementById('agent-prompt-input');
  const prompt = input.value.trim();
  if (!prompt) return;

  input.value = '';
  const btn = document.getElementById('btn-agent-send');
  btn.disabled = true;

  appendAgentLog(`\n[USER PROMPT]: ${prompt}`, true);
  setSystemStatus('● AGENT REASONING...', true);

  try {
    const res = await fetch('/api/agent/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ prompt })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error);

    const evtSource = new EventSource(`/api/generator/stream/${data.job_id}`);
    evtSource.addEventListener('log', (e) => {
      const d = JSON.parse(e.data);
      appendAgentLog(d.message);
    });

    evtSource.addEventListener('completed', (e) => {
      const d = JSON.parse(e.data);
      appendAgentLog(`\n[AGENT RESPONSE]:\n${d.response}\n`);
      setSystemStatus('● SYSTEM READY');
      btn.disabled = false;
    });

    evtSource.addEventListener('error', () => {
      appendAgentLog(`[!] Agent execution error.`);
      setSystemStatus('● AGENT ERROR');
      btn.disabled = false;
    });

    evtSource.addEventListener('close', () => evtSource.close());
  } catch (err) {
    appendAgentLog(`[!] Agent failed: ${err.message}`);
    setSystemStatus('● AGENT ERROR');
    btn.disabled = false;
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
      alert('Configuration updated and saved to .env file successfully.');
      loadConfigBadge();
    }
  } catch (err) {
    alert(`Save failed: ${err.message}`);
  }
}
