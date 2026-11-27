/* TimeStudyPro Cloud Sync add-on for V260. Uses only Supabase publishable key + authenticated RLS. */
(function () {
  'use strict';
  const SUPABASE_URL = 'https://utenvzaylwuiuqxmmwjj.supabase.co';
  const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_uZ1NvovzyQGlj9ngQMV8Bg_ndzOvFMn';
  const BUCKET = 'timestudypro-files';
  let client = null;
  let currentUser = null;
  const el = (tag, props = {}, html = '') => { const n = document.createElement(tag); Object.assign(n, props); if (html) n.innerHTML = html; return n; };
  function getClient() {
    if (client) return client;
    if (!window.supabase || !window.supabase.createClient) throw new Error('Supabase client មិនទាន់ Load បានទេ។ សូមពិនិត្យ Internet រួច Refresh។');
    client = window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true } });
    return client;
  }
  window.tspCloudIsSignedIn = () => !!currentUser;
  window.tspCloudUploadFile = async (metadata, file) => {
    const sb = getClient();
    const { data: authData } = await sb.auth.getUser();
    const user = authData && authData.user;
    if (!user) throw new Error('សូម Cloud Login ជាមុនសិន។');
    const safeName = String(file.name || 'excel.xlsx').replace(/[^\w.\-\u00A0-\uFFFF]+/g, '_');
    const path = `${user.id}/${Date.now()}_${safeName}`;
    const upload = await sb.storage.from(BUCKET).upload(path, file, { contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', upsert: false });
    if (upload.error) throw upload.error;
    const row = { user_id: user.id, file_name: file.name, storage_path: path, mime_type: file.type || null, size_bytes: file.size || 0, metadata: { ...metadata, cloudId: undefined, cloudStoragePath: undefined } };
    const inserted = await sb.from('timestudypro_files').insert(row).select('id,storage_path').single();
    if (inserted.error) {
      await sb.storage.from(BUCKET).remove([path]).catch(() => {});
      throw inserted.error;
    }
    return inserted.data;
  };
  async function syncFromCloud() {
    const sb = getClient();
    const { data: authData, error: authError } = await sb.auth.getUser();
    if (authError) throw authError;
    currentUser = authData.user || null;
    if (!currentUser) throw new Error('សូម Login ទៅ Cloud ជាមុនសិន។');
    const { data, error } = await sb.from('timestudypro_files').select('id,file_name,storage_path,mime_type,size_bytes,created_at,metadata').order('created_at', { ascending: false });
    if (error) throw error;
    const added = typeof window.tspCloudMergeFiles === 'function' ? await window.tspCloudMergeFiles(data || []) : 0;
    status.textContent = `បានភ្ជាប់ Cloud: ${currentUser.email || 'គណនី Cloud'} · ទាញយកទិន្នន័យ ${data.length} · បន្ថែមថ្មី ${added}`;
    status.style.color = '#15803d';
    loginButton.textContent = 'Cloud Sync';
    logoutButton.style.display = 'inline-flex';
    return added;
  }
  const style = el('style', {}, `#tsp-cloud-box{position:fixed;right:16px;bottom:18px;z-index:99999;font:13px/1.45 Arial,sans-serif;color:#e5eefb}#tsp-cloud-box button{border:0;border-radius:9px;padding:9px 12px;cursor:pointer;font-weight:700}#tsp-cloud-toggle{background:#153a65;color:#fff;box-shadow:0 4px 16px #0003}#tsp-cloud-panel{display:none;width:min(340px,calc(100vw - 32px));margin-top:8px;padding:14px;background:#0f1f36;border:1px solid #355174;border-radius:14px;box-shadow:0 12px 35px #0006}#tsp-cloud-panel label{display:block;margin:9px 0 4px;color:#cbd5e1}#tsp-cloud-panel input{box-sizing:border-box;width:100%;padding:10px;border:1px solid #49627f;border-radius:8px;background:#f8fafc;color:#111827}#tsp-cloud-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}#tsp-cloud-actions button{background:#2563eb;color:#fff}#tsp-cloud-actions #tsp-cloud-logout{background:#475569}#tsp-cloud-status{font-size:12px;margin-top:10px;overflow-wrap:anywhere;color:#cbd5e1}#tsp-cloud-box .tsp-cloud-small{font-size:11px;color:#94a3b8;margin-top:8px}`);
  document.head.appendChild(style);
  const box = el('div', { id: 'tsp-cloud-box' });
  const toggle = el('button', { id: 'tsp-cloud-toggle', type: 'button' }, '☁ Cloud Login');
  const panel = el('div', { id: 'tsp-cloud-panel' });
  panel.appendChild(el('div', {}, '<strong>TimeStudyPro Cloud</strong><div class="tsp-cloud-small">ចូលគណនី Supabase ដូចគ្នាលើទូរស័ព្ទ និងកុំព្យូទ័រ</div>'));
  panel.appendChild(el('label', { htmlFor: 'tsp-cloud-email' }, 'Gmail / Email'));
  const email = el('input', { id: 'tsp-cloud-email', type: 'email', autocomplete: 'username', placeholder: 'name@gmail.com' }); panel.appendChild(email);
  panel.appendChild(el('label', { htmlFor: 'tsp-cloud-password' }, 'Password'));
  const password = el('input', { id: 'tsp-cloud-password', type: 'password', autocomplete: 'current-password', placeholder: 'Cloud account password' }); panel.appendChild(password);
  const actions = el('div', { id: 'tsp-cloud-actions' });
  const loginButton = el('button', { type: 'button' }, 'Login & Sync');
  const syncButton = el('button', { type: 'button' }, 'Sync ឥឡូវ');
  const logoutButton = el('button', { id: 'tsp-cloud-logout', type: 'button' }, 'Logout'); logoutButton.style.display = 'none';
  actions.append(loginButton, syncButton, logoutButton); panel.appendChild(actions);
  const status = el('div', { id: 'tsp-cloud-status' }, 'មិនទាន់ភ្ជាប់ Cloud'); panel.appendChild(status);
  box.append(toggle, panel); document.body.appendChild(box);
  toggle.addEventListener('click', () => { panel.style.display = panel.style.display === 'block' ? 'none' : 'block'; });
  loginButton.addEventListener('click', async () => {
    loginButton.disabled = true; status.textContent = 'កំពុង Login...'; status.style.color = '#cbd5e1';
    try {
      const sb = getClient();
      const { data, error } = await sb.auth.signInWithPassword({ email: email.value.trim(), password: password.value });
      if (error) throw error;
      currentUser = data.user;
      await syncFromCloud();
    } catch (e) { status.textContent = 'Cloud Login/Sync មិនជោគជ័យ: ' + (e.message || e); status.style.color = '#fca5a5'; }
    finally { loginButton.disabled = false; }
  });
  syncButton.addEventListener('click', async () => {
    syncButton.disabled = true;
    try { await syncFromCloud(); }
    catch (e) { status.textContent = 'Sync មិនជោគជ័យ: ' + (e.message || e); status.style.color = '#fca5a5'; }
    finally { syncButton.disabled = false; }
  });
  logoutButton.addEventListener('click', async () => {
    try { await getClient().auth.signOut(); currentUser = null; status.textContent = 'បាន Logout ពី Cloud'; status.style.color = '#cbd5e1'; logoutButton.style.display = 'none'; loginButton.textContent = 'Login & Sync'; }
    catch (e) { status.textContent = 'Logout មិនជោគជ័យ: ' + (e.message || e); }
  });
  // Restore an existing session and sync quietly when the app opens.
  try {
    getClient().auth.getSession().then(({ data, error }) => {
      if (!error && data && data.session) { currentUser = data.session.user; syncFromCloud().catch(e => { status.textContent = 'Cloud Sync ត្រូវការពិនិត្យ: ' + (e.message || e); status.style.color = '#fca5a5'; }); }
    }).catch(() => {});
  } catch (e) { status.textContent = 'Supabase library មិនទាន់ Load បាន។ អ្នកនៅតែអាចប្រើទិន្នន័យក្នុង Browser បាន។'; }
})();
