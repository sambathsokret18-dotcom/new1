/* TimeStudyPro Pre B/L file cabinet. Stores original files in browser IndexedDB. */
(function () {
  'use strict';
  if (window.__TSP_PREBL_V1__) return;
  window.__TSP_PREBL_V1__ = true;
  const DB_NAME = 'TimeStudyProPreBLStorage';
  const STORE = 'files';
  const input = document.getElementById('prebl-file-input');
  const list = document.getElementById('prebl-file-list');
  const search = document.getElementById('prebl-search');
  const count = document.getElementById('prebl-count');
  const gridButton = document.getElementById('prebl-grid-view');
  const listButton = document.getElementById('prebl-list-view');
  let viewMode = 'grid';
  if (!input || !list) return;
  let db;
  let files = [];
  const openDB = () => new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE, {keyPath:'id'}); };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Cannot open local file storage'));
  });
  const txRequest = (mode, fn) => new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode); const req = fn(tx.objectStore(STORE));
    if (req) { req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); }
    else { tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); }
  });
  const fmtSize = bytes => bytes < 1024 ? `${bytes} B` : bytes < 1048576 ? `${(bytes/1024).toFixed(1)} KB` : `${(bytes/1048576).toFixed(2)} MB`;
  const fmtDate = value => { try { return new Date(value).toLocaleString(); } catch (_) { return ''; } };
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const iconFor = name => { const ext = (name.split('.').pop() || '').toLowerCase(); return ['xls','xlsx','csv'].includes(ext) ? 'fa-file-excel' : ['pdf'].includes(ext) ? 'fa-file-pdf' : ['doc','docx'].includes(ext) ? 'fa-file-word' : ['zip','rar','7z'].includes(ext) ? 'fa-file-zipper' : ['png','jpg','jpeg','gif','webp'].includes(ext) ? 'fa-file-image' : 'fa-file-lines'; };
  async function refresh() {
    try { files = await txRequest('readonly', store => store.getAll()); render(); }
    catch (error) { list.innerHTML = `<div class="prebl-empty">មិនអាចអានឯកសារបាន៖ ${esc(error.message)}</div>`; }
  }
  function render() {
    const q = (search?.value || '').trim().toLowerCase();
    const shown = files.filter(file => file.name.toLowerCase().includes(q)).sort((a,b) => b.savedAt-a.savedAt);
    if (count) count.textContent = `${shown.length} / ${files.length} files`;
    list.classList.toggle('prebl-list-mode', viewMode === 'list');
    if (gridButton) { gridButton.classList.toggle('active', viewMode === 'grid'); gridButton.setAttribute('aria-pressed', String(viewMode === 'grid')); }
    if (listButton) { listButton.classList.toggle('active', viewMode === 'list'); listButton.setAttribute('aria-pressed', String(viewMode === 'list')); }
    if (!shown.length) { list.innerHTML = `<div class="prebl-empty">${files.length ? 'រកមិនឃើញឯកសារដែលត្រូវនឹងពាក្យស្វែងរកទេ។' : 'មិនទាន់មានឯកសារទេ។ ចុច Add Files ដើម្បីរក្សាទុកឯកសារ។'}</div>`; return; }
    list.innerHTML = shown.map(file => `<article class="prebl-file-card"><div class="prebl-file-top"><div class="prebl-file-icon"><i class="fa-solid ${iconFor(file.name)}"></i></div><div style="min-width:0;flex:1"><div class="prebl-file-name" title="${esc(file.name)}">${esc(file.name)}</div><div class="prebl-file-meta">${fmtSize(file.size)} · ${fmtDate(file.savedAt)}</div></div></div><div class="prebl-file-actions"><button type="button" class="prebl-download" data-prebl-download="${esc(file.id)}"><i class="fa-solid fa-download"></i> Download</button><button type="button" class="prebl-delete" data-prebl-delete="${esc(file.id)}"><i class="fa-solid fa-trash"></i> Delete</button></div></article>`).join('');
  }
  input.addEventListener('change', async () => {
    const selected = Array.from(input.files || []); if (!selected.length) return;
    try {
      const existingNames = new Set(files.map(item => item.name.trim().toLocaleLowerCase()));
      const batchNames = new Set();
      const duplicates = [];
      const accepted = [];
      for (const file of selected) {
        const normalizedName = file.name.trim().toLocaleLowerCase();
        if (existingNames.has(normalizedName) || batchNames.has(normalizedName)) duplicates.push(file.name);
        else { batchNames.add(normalizedName); accepted.push(file); }
      }
      for (const file of accepted) {
        await txRequest('readwrite', store => store.put({id: `${Date.now()}_${Math.random().toString(36).slice(2)}_${file.name}`, name:file.name, type:file.type || 'application/octet-stream', size:file.size, savedAt:Date.now(), blob:file}));
      }
      await refresh();
      if (duplicates.length) {
        const names = [...new Set(duplicates)].slice(0, 8).join('、');
        const more = duplicates.length > 8 ? ` និង ${duplicates.length - 8} ឯកសារផ្សេងទៀត` : '';
        const message = `មិនបាន Upload ឯកសារឈ្មោះដូចគ្នា៖ ${names}${more}`;
        if (window.Swal) await Swal.fire({icon:'warning',title:'មានឈ្មោះឯកសារស្ទួន',text:message,confirmButtonText:'យល់ព្រម'});
        else alert(message);
      }
      if (accepted.length && window.Swal) Swal.fire({icon:'success',title:'បានរក្សាទុកឯកសារ',text:`បានរក្សាទុក ${accepted.length} ឯកសារ`,timer:1700,showConfirmButton:false});
    } catch (error) { alert('រក្សាទុកឯកសារមិនបាន៖ ' + (error.message || error)); }
    input.value = '';
  });
  list.addEventListener('click', async event => {
    const download = event.target.closest('[data-prebl-download]');
    const remove = event.target.closest('[data-prebl-delete]');
    if (download) {
      const file = files.find(item => item.id === download.dataset.preblDownload); if (!file) return;
      const url = URL.createObjectURL(file.blob); const a = document.createElement('a'); a.href = url; a.download = file.name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500);
    } else if (remove) {
      const file = files.find(item => item.id === remove.dataset.preblDelete); if (!file) return;
      let ok = window.confirm(`តើអ្នកប្រាកដថាចង់លុបឯកសារ “${file.name}” មែនទេ?`); if (!ok) return;
      try { await txRequest('readwrite', store => store.delete(file.id)); await refresh(); }
      catch (error) { alert('លុបឯកសារមិនបាន៖ ' + (error.message || error)); }
    }
  });
  if (search) search.addEventListener('input', render);
  if (gridButton) gridButton.addEventListener('click', () => { viewMode = 'grid'; render(); });
  if (listButton) listButton.addEventListener('click', () => { viewMode = 'list'; render(); });
  openDB().then(async value => { db = value; await refresh(); }).catch(error => { list.innerHTML = `<div class="prebl-empty">មិនអាចបើកកន្លែងរក្សាទុកឯកសារបាន៖ ${esc(error.message)}</div>`; });
  document.querySelectorAll('[data-page="prebl-page"], [data-mobile-action="page"][data-page="prebl-page"]').forEach(link => link.addEventListener('click', () => { if (db) refresh(); }));
})();
