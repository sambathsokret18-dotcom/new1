/* TimeStudyPro VDO storage: preview the selected video in a modal, then require Style and Process before saving. */
(function () {
  'use strict';
  if (window.__TSP_VDO_STORAGE_V3__) return;
  window.__TSP_VDO_STORAGE_V3__ = true;
  const DB_NAME = 'TimeStudyProVDOStorage';
  const STORE = 'videos';
  const input = document.getElementById('vdo-file-input');
  const list = document.getElementById('vdo-file-list');
  const search = document.getElementById('vdo-search');
  const processSearch = document.getElementById('vdo-process-search');
  const count = document.getElementById('vdo-count');
  const totalStyles = document.getElementById('vdo-total-styles');
  const totalVideos = document.getElementById('vdo-total-videos');
  const gridButton = document.getElementById('vdo-grid-view');
  const listButton = document.getElementById('vdo-list-view');
  const pagination = document.getElementById('vdo-pagination');
  const PAGE_SIZE = 12;
  let currentPage = 1;
  const selectedIds = new Set();
  let visiblePageIds = [];
  const modal = document.getElementById('vdo-upload-modal');
  const modalPreview = document.getElementById('vdo-modal-preview');
  const modalFilename = document.getElementById('vdo-modal-filename');
  const modalClose = document.getElementById('vdo-modal-close');
  const modalCancel = document.getElementById('vdo-modal-cancel');
  const modalSave = document.getElementById('vdo-modal-save');
  const styleInput = document.getElementById('vdo-style-input');
  const processInput = document.getElementById('vdo-process-input');
  if (!input || !list || !modal || !modalSave || !styleInput || !processInput) return;
  let db;
  let videos = [];
  let viewMode = 'grid';
  let pendingFile = null;
  let pendingPreviewUrl = null;
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const norm = value => String(value ?? '').trim().toLocaleLowerCase();
  const fmtSize = bytes => bytes < 1024 ? `${bytes} B` : bytes < 1048576 ? `${(bytes/1024).toFixed(1)} KB` : bytes < 1073741824 ? `${(bytes/1048576).toFixed(2)} MB` : `${(bytes/1073741824).toFixed(2)} GB`;
  const fmtDate = value => { try { return new Date(value).toLocaleString(); } catch (_) { return ''; } };
  const isVideo = file => !!file && ((file.type || '').startsWith('video/') || /\.(mp4|mov|avi|mkv|webm|mpeg|mpg|3gp|m4v|wmv)$/i.test(file.name));
  const openDB = () => new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE, {keyPath:'id'}); };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Cannot open local video storage'));
  });
  const txRequest = (mode, fn) => new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode); const req = fn(tx.objectStore(STORE));
    if (req) { req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); }
    else { tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); }
  });
  async function refresh() {
    try { videos = await txRequest('readonly', store => store.getAll()); render(); }
    catch (error) { list.innerHTML = `<div class="prebl-empty">មិនអាចអាន VDO បាន៖ ${esc(error.message)}</div>`; }
  }
  function render() {
    const q = norm(search?.value || '');
    const processQ = norm(processSearch?.value || '');
    const shown = videos.filter(video =>
      (!q || [video.name, video.style].some(value => norm(value).includes(q))) &&
      (!processQ || norm(video.process).includes(processQ))
    ).sort((a,b) => b.savedAt-a.savedAt);
    if (count) count.textContent = `${shown.length} / ${videos.length} VDO`;
    if (totalStyles) totalStyles.textContent = String(new Set(videos.map(video => norm(video.style)).filter(Boolean)).size);
    if (totalVideos) totalVideos.textContent = String(videos.length);
    list.classList.toggle('prebl-list-mode', viewMode === 'list');
    if (gridButton) { gridButton.classList.toggle('active', viewMode === 'grid'); gridButton.setAttribute('aria-pressed', String(viewMode === 'grid')); }
    if (listButton) { listButton.classList.toggle('active', viewMode === 'list'); listButton.setAttribute('aria-pressed', String(viewMode === 'list')); }
    const pageCount = Math.max(1, Math.ceil(shown.length / PAGE_SIZE));
    currentPage = Math.min(Math.max(1, currentPage), pageCount);
    if (!shown.length) {
      visiblePageIds = [];
      const selectPage = document.getElementById('vdo-select-page'); if (selectPage) { selectPage.checked = false; selectPage.indeterminate = false; }
      const selectedCount = document.getElementById('vdo-selected-count'); if (selectedCount) selectedCount.textContent = `${selectedIds.size} selected`;
      const deleteSelected = document.getElementById('vdo-delete-selected'); if (deleteSelected) deleteSelected.disabled = selectedIds.size === 0;
      list.innerHTML = `<div class="prebl-empty">${videos.length ? 'រកមិនឃើញ VDO តាមឈ្មោះ, Style ឬ Process ទេ។' : 'មិនទាន់មាន VDO ទេ។ ចុច Upload VDO ដើម្បីបន្ថែម។'}</div>`;
      if (pagination) { pagination.hidden = true; pagination.innerHTML = ''; }
      return;
    }
    const start = (currentPage - 1) * PAGE_SIZE;
    const pageVideos = shown.slice(start, start + PAGE_SIZE);
    visiblePageIds = pageVideos.map(video => video.id);
    list.addEventListener('change', event => {
    const select = event.target.closest('[data-vdo-select]');
    if (!select) return;
    if (select.checked) selectedIds.add(select.dataset.vdoSelect); else selectedIds.delete(select.dataset.vdoSelect);
    render();
  });
  const selectPage = document.getElementById('vdo-select-page');
    if (selectPage) { selectPage.checked = visiblePageIds.length > 0 && visiblePageIds.every(id => selectedIds.has(id)); selectPage.indeterminate = visiblePageIds.some(id => selectedIds.has(id)) && !selectPage.checked; }
    const selectedCount = document.getElementById('vdo-selected-count');
    if (selectedCount) selectedCount.textContent = `${selectedIds.size} selected`;
    const deleteSelected = document.getElementById('vdo-delete-selected');
    if (deleteSelected) deleteSelected.disabled = selectedIds.size === 0;
    list.innerHTML = pageVideos.map(video => `<article class="prebl-file-card"><div class="vdo-select-row"><label><input type="checkbox" data-vdo-select="${esc(video.id)}" ${selectedIds.has(video.id) ? 'checked' : ''}><span>Select VDO</span></label></div><div class="prebl-file-top"><div class="prebl-file-icon vdo-file-icon"><i class="fa-solid fa-film"></i></div><div style="min-width:0;flex:1"><div class="prebl-file-name" title="${esc(video.name)}">${esc(video.name)}</div><div class="prebl-file-meta">${fmtSize(video.size)} · ${fmtDate(video.savedAt)}</div></div></div><div class="vdo-file-tags"><span class="vdo-file-tag"><b>Style:</b> ${esc(video.style || '—')}</span><span class="vdo-file-tag"><b>Process:</b> ${esc(video.process || '—')}</span></div><div class="prebl-file-actions"><button type="button" class="vdo-preview" data-vdo-preview="${esc(video.id)}"><i class="fa-solid fa-play"></i> មើល</button><button type="button" class="prebl-download" data-vdo-download="${esc(video.id)}"><i class="fa-solid fa-download"></i> Download</button><button type="button" class="prebl-delete" data-vdo-delete="${esc(video.id)}"><i class="fa-solid fa-trash"></i> Delete</button></div></article>`).join('');
    if (pagination) {
      const windowStart = Math.max(1, Math.min(currentPage - 2, pageCount - 4));
      const windowEnd = Math.min(pageCount, windowStart + 4);
      const pageButtons = [];
      pageButtons.push(`<button type="button" data-vdo-page="${currentPage - 1}" ${currentPage === 1 ? 'disabled' : ''} aria-label="Previous page">‹ Prev</button>`);
      if (windowStart > 1) pageButtons.push(`<button type="button" data-vdo-page="1">1</button>${windowStart > 2 ? '<span class="vdo-page-ellipsis">…</span>' : ''}`);
      for (let page = windowStart; page <= windowEnd; page++) pageButtons.push(`<button type="button" data-vdo-page="${page}" class="${page === currentPage ? 'active' : ''}" ${page === currentPage ? 'aria-current="page"' : ''}>${page}</button>`);
      if (windowEnd < pageCount) pageButtons.push(`${windowEnd < pageCount - 1 ? '<span class="vdo-page-ellipsis">…</span>' : ''}<button type="button" data-vdo-page="${pageCount}">${pageCount}</button>`);
      pageButtons.push(`<button type="button" data-vdo-page="${currentPage + 1}" ${currentPage === pageCount ? 'disabled' : ''} aria-label="Next page">Next ›</button>`);
      pagination.innerHTML = `<span class="vdo-page-summary">${start + 1}–${Math.min(start + PAGE_SIZE, shown.length)} of ${shown.length}</span><div class="vdo-page-buttons">${pageButtons.join('')}</div>`;
      pagination.hidden = pageCount <= 1;
    }
  }
  function releasePendingPreview() {
    if (pendingPreviewUrl) URL.revokeObjectURL(pendingPreviewUrl);
    pendingPreviewUrl = null;
    if (modalPreview) { modalPreview.pause(); modalPreview.removeAttribute('src'); modalPreview.load(); }
  }
  function closeModal() {
    modal.hidden = true;
    releasePendingPreview();
    pendingFile = null;
    styleInput.value = '';
    processInput.value = '';
    modalSave.disabled = true;
    modalFilename.textContent = '';
    input.value = '';
  }
  function updateSaveState() { modalSave.disabled = !(pendingFile && styleInput.value.trim() && processInput.value.trim()); }
  function openModal(file) {
    if (!isVideo(file)) { alert('សូមជ្រើសរើសឯកសារ VDO ដែលត្រឹមត្រូវ។'); input.value = ''; return; }
    pendingFile = file;
    releasePendingPreview();
    pendingPreviewUrl = URL.createObjectURL(file);
    modalPreview.src = pendingPreviewUrl;
    modalFilename.textContent = `${file.name} · ${fmtSize(file.size)}`;
    styleInput.value = '';
    processInput.value = '';
    modalSave.disabled = true;
    modal.hidden = false;
    setTimeout(() => styleInput.focus(), 40);
  }
  input.addEventListener('change', () => {
    const file = input.files && input.files[0];
    if (file) openModal(file);
  });
  styleInput.addEventListener('input', updateSaveState);
  processInput.addEventListener('input', updateSaveState);
  modalClose.addEventListener('click', closeModal);
  modalCancel.addEventListener('click', closeModal);
  modal.addEventListener('click', event => { if (event.target === modal) closeModal(); });
  document.addEventListener('keydown', event => { if (event.key === 'Escape' && !modal.hidden) closeModal(); });
  modalSave.addEventListener('click', async () => {
    const style = styleInput.value.trim();
    const process = processInput.value.trim();
    if (!pendingFile || !style || !process) { updateSaveState(); alert('សូមបំពេញឈ្មោះ Style និង Process ទាំងពីរ មុន Save។'); return; }
    const duplicate = videos.some(video => norm(video.name) === norm(pendingFile.name) && norm(video.style) === norm(style));
    if (duplicate) { alert(`មាន VDO ឈ្មោះ “${pendingFile.name}” សម្រាប់ Style “${style}” រួចហើយ។ មិនអាច Save ស្ទួនបានទេ។`); return; }
    // Keep the Save label unchanged; prevent double-clicks without showing a loading indicator.
    modalSave.disabled = true;
    try {
      await txRequest('readwrite', store => store.put({id:`${Date.now()}_${Math.random().toString(36).slice(2)}_${pendingFile.name}`, name:pendingFile.name, type:pendingFile.type || 'video/mp4', size:pendingFile.size, style, process, savedAt:Date.now(), blob:pendingFile}));
      currentPage = 1;
      await refresh();
      closeModal();
      if (window.Swal) Swal.fire({icon:'success',title:'បានរក្សាទុក VDO',text:`${style} · ${process}`,timer:1800,showConfirmButton:false});
    } catch (error) {
      alert('រក្សាទុក VDO មិនបាន៖ ' + (error.message || error) + '\nសូមពិនិត្យទំហំ Storage របស់ Browser។');
      updateSaveState();
    } finally {
      modalSave.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Save VDO';
    }
  });
  list.addEventListener('click', async event => {
    const preview = event.target.closest('[data-vdo-preview]');
    const download = event.target.closest('[data-vdo-download]');
    const remove = event.target.closest('[data-vdo-delete]');
    const video = videos.find(item => item.id === (preview?.dataset.vdoPreview || download?.dataset.vdoDownload || remove?.dataset.vdoDelete));
    if (!video) return;
    if (preview) {
      const url = URL.createObjectURL(video.blob);
      const win = window.open('', '_blank');
      if (!win) { URL.revokeObjectURL(url); alert('Browser បានបិទ Popup។ សូមអនុញ្ញាត Popup រួចសាកម្ដងទៀត។'); return; }
      win.document.title = video.name;
      win.document.body.style.cssText = 'margin:0;background:#0f172a;display:grid;place-items:center;min-height:100vh;color:white;font-family:Arial,sans-serif';
      const player = win.document.createElement('video'); player.controls = true; player.autoplay = true; player.src = url; player.style.cssText = 'width:100%;max-height:100vh';
      const heading = win.document.createElement('div'); heading.textContent = `${video.name} · Style: ${video.style || '—'} · Process: ${video.process || '—'}`; heading.style.cssText = 'position:fixed;top:8px;left:12px;color:white;background:#0009;padding:6px 10px;border-radius:6px;font-size:13px;max-width:90vw';
      win.document.body.append(player, heading); win.addEventListener('beforeunload', () => URL.revokeObjectURL(url));
    } else if (download) {
      const url = URL.createObjectURL(video.blob); const a = document.createElement('a'); a.href = url; a.download = video.name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000);
    } else if (remove) {
      if (!window.confirm(`តើអ្នកប្រាកដថាចង់លុប VDO “${video.name}” (Style: ${video.style || '—'}, Process: ${video.process || '—'}) មែនទេ?`)) return;
      try { await txRequest('readwrite', store => store.delete(video.id)); selectedIds.delete(video.id); await refresh(); }
      catch (error) { alert('លុប VDO មិនបាន៖ ' + (error.message || error)); }
    }
  });
  const selectPage = document.getElementById('vdo-select-page');
  if (selectPage) selectPage.addEventListener('change', () => {
    visiblePageIds.forEach(id => selectPage.checked ? selectedIds.add(id) : selectedIds.delete(id));
    render();
  });
  const deleteSelected = document.getElementById('vdo-delete-selected');
  if (deleteSelected) deleteSelected.addEventListener('click', async () => {
    const ids = Array.from(selectedIds).filter(id => videos.some(video => video.id === id));
    if (!ids.length) { alert('សូមជ្រើសរើស VDO ដែលចង់លុបជាមុន។'); return; }
    if (!window.confirm(`តើអ្នកប្រាកដថាចង់លុប VDO ចំនួន ${ids.length} ក្នុងពេលតែមួយមែនទេ? មិនអាចត្រឡប់វិញបានទេ។`)) return;
    deleteSelected.disabled = true;
    try {
      await new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, 'readwrite'); const store = tx.objectStore(STORE);
        ids.forEach(id => store.delete(id));
        tx.oncomplete = resolve; tx.onerror = () => reject(tx.error || new Error('Delete failed')); tx.onabort = () => reject(tx.error || new Error('Delete aborted'));
      });
      ids.forEach(id => selectedIds.delete(id));
      await refresh();
      if (window.Swal) Swal.fire({icon:'success',title:'បានលុប VDO រួចរាល់',text:`បានលុប ${ids.length} VDO`,timer:1800,showConfirmButton:false});
    } catch (error) { alert('លុប VDO ច្រើនមិនបាន៖ ' + (error.message || error)); }
    finally { render(); }
  });
  if (search) search.addEventListener('input', () => { currentPage = 1; render(); });
  if (processSearch) processSearch.addEventListener('input', () => { currentPage = 1; render(); });
  if (pagination) pagination.addEventListener('click', event => { const button = event.target.closest('[data-vdo-page]'); if (!button || button.disabled) return; currentPage = Number(button.dataset.vdoPage) || 1; render(); });
  if (gridButton) gridButton.addEventListener('click', () => { viewMode = 'grid'; render(); });
  if (listButton) listButton.addEventListener('click', () => { viewMode = 'list'; render(); });
  async function init() {
    try { if (window.tspLargeStorageReady) await window.tspLargeStorageReady; } catch (_) {}
    try { db = await openDB(); await refresh(); }
    catch (error) { list.innerHTML = `<div class="prebl-empty">មិនអាចបើកកន្លែងរក្សាទុក VDO បាន៖ ${esc(error.message)}</div>`; }
  }
  init();
  document.querySelectorAll('[data-page="video-storage-page"], [data-mobile-action="page"][data-page="video-storage-page"]').forEach(link => link.addEventListener('click', () => { if (db) refresh(); }));
})();
