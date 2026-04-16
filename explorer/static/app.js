/**
 * Horde Email Explorer — frontend logic
 * Pure vanilla JS, no framework dependencies.
 */

'use strict';

// ── State ─────────────────────────────────────────────────────────────────
const state = {
  currentFolderId: null,
  currentFolderName: null,
  currentView: 'folder',   // 'folder' | 'search' | 'bookmarks'
  searchQuery: '',
  page: 1,
  perPage: 50,
  sortField: 'date_sent',
  sortOrder: 'desc',
  filters: {},
  currentEmailId: null,
  bodyMode: 'html',         // 'html' | 'plain'
  theme: localStorage.getItem('theme') || 'light',
};

// ── Helpers ───────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const esc = s => (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

function fmt_date(iso) {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    const now = new Date();
    const sameYear = d.getFullYear() === now.getFullYear();
    const sameDay  = d.toDateString() === now.toDateString();
    if (sameDay)  return d.toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'});
    if (sameYear) return d.toLocaleDateString([], {month:'short',day:'numeric'});
    return d.toLocaleDateString([], {year:'2-digit',month:'short',day:'numeric'});
  } catch { return iso; }
}

function fmt_size(bytes) {
  if (bytes == null) return '';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024*1024) return (bytes/1024).toFixed(1) + ' KB';
  return (bytes/1024/1024).toFixed(1) + ' MB';
}

async function apiFetch(url) {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  return resp.json();
}

// ── Theme ─────────────────────────────────────────────────────────────────
function applyTheme() {
  document.documentElement.setAttribute('data-theme', state.theme);
  localStorage.setItem('theme', state.theme);
  $('theme-toggle').textContent = state.theme === 'dark' ? '☀' : '☾';
  // Reload iframe body so its background matches
  const iframe = $('body-iframe');
  if (iframe.src) {
    const src = iframe.src;
    iframe.src = '';
    iframe.src = src;
  }
}

$('theme-toggle').addEventListener('click', () => {
  state.theme = state.theme === 'dark' ? 'light' : 'dark';
  applyTheme();
});
applyTheme();

// ── Stats ─────────────────────────────────────────────────────────────────
async function loadStats() {
  try {
    const s = await apiFetch('/api/stats');
    $('stats-bar').textContent =
      `${s.total_emails.toLocaleString()} emails · ${s.total_folders} folders`;
  } catch {}
}

// ── Folders ───────────────────────────────────────────────────────────────
async function loadFolders() {
  const list = $('folder-list');
  list.innerHTML = '<li style="padding:8px 14px;font-size:12px;opacity:.5">Loading…</li>';
  try {
    const folders = await apiFetch('/api/folders');
    list.innerHTML = '';
    folders.forEach(f => {
      const li = document.createElement('li');
      li.className = 'folder-item';
      li.dataset.folderId = f.id;
      li.dataset.folderName = f.name;
      li.innerHTML = `<span class="folder-icon">📁</span>
        <span class="folder-label">${esc(f.name)}</span>
        <span class="folder-count">${f.total}</span>`;
      li.addEventListener('click', () => selectFolder(f.id, f.name));
      list.appendChild(li);
    });
    // Auto-select first folder
    if (folders.length > 0) selectFolder(folders[0].id, folders[0].name);
  } catch (e) {
    list.innerHTML = `<li style="padding:8px 14px;color:#e74c3c">Error: ${esc(e.message)}</li>`;
  }
}

function selectFolder(id, name) {
  state.currentFolderId = id;
  state.currentFolderName = name;
  state.currentView = 'folder';
  state.page = 1;
  // Update active state
  document.querySelectorAll('.folder-item').forEach(el => {
    el.classList.toggle('active', el.dataset.folderId == id);
  });
  $('search-input').value = '';
  clearFilters();
  loadEmailList();
}

// ── Email list ────────────────────────────────────────────────────────────
async function loadEmailList() {
  const container = $('email-list');
  const status = $('list-status');
  container.innerHTML = '';
  status.textContent = 'Loading…';

  let url;
  if (state.currentView === 'search' && state.searchQuery) {
    const params = new URLSearchParams({
      q: state.searchQuery,
      page: state.page,
      per_page: state.perPage,
    });
    if (state.currentFolderId) params.set('folder_id', state.currentFolderId);
    url = `/api/search?${params}`;
  } else if (state.currentView === 'bookmarks') {
    renderBookmarks();
    return;
  } else {
    const params = new URLSearchParams({
      page: state.page,
      per_page: state.perPage,
      sort: state.sortField,
      order: state.sortOrder,
    });
    if (state.currentFolderId) params.set('folder_id', state.currentFolderId);
    Object.entries(state.filters).forEach(([k, v]) => { if (v) params.set(k, v); });
    url = `/api/emails?${params}`;
  }

  try {
    const data = await apiFetch(url);
    const emails = data.emails || [];
    status.textContent = '';

    if (emails.length === 0) {
      status.textContent = 'No emails found.';
      updatePagination(0);
      return;
    }

    emails.forEach(e => {
      const li = document.createElement('li');
      li.className = 'email-item' + (e.bookmarked ? ' bookmarked' : '');
      li.dataset.id = e.id;
      li.innerHTML = `
        <span class="ei-from">${esc(e.sender || '(unknown)')}</span>
        <span class="ei-date">${fmt_date(e.date_sent)}</span>
        <span class="ei-subject">${esc(e.subject || '(no subject)')}</span>
        <span class="ei-folder">${esc(e.folder_name || '')}</span>
        ${e.snippet ? `<span class="ei-snippet">${e.snippet}</span>` : ''}
        ${e.has_attachments ? '<span class="ei-att">📎</span>' : ''}
      `;
      li.addEventListener('click', () => openEmail(e.id));
      container.appendChild(li);
    });

    updatePagination(data.total);
  } catch (err) {
    status.textContent = 'Error loading emails: ' + err.message;
  }
}

function updatePagination(total) {
  const totalPages = Math.ceil(total / state.perPage);
  $('page-prev').disabled = state.page <= 1;
  $('page-next').disabled = state.page >= totalPages;
  $('page-info').textContent = total
    ? `${((state.page-1)*state.perPage)+1}–${Math.min(state.page*state.perPage,total)} of ${total.toLocaleString()}`
    : '';
}

$('page-prev').addEventListener('click', () => { state.page--; loadEmailList(); });
$('page-next').addEventListener('click', () => { state.page++; loadEmailList(); });

// ── Sort ──────────────────────────────────────────────────────────────────
$('sort-field').addEventListener('change', e => {
  state.sortField = e.target.value; state.page = 1; loadEmailList();
});
$('sort-order').addEventListener('change', e => {
  state.sortOrder = e.target.value; state.page = 1; loadEmailList();
});

// ── Filter ────────────────────────────────────────────────────────────────
$('filter-apply').addEventListener('click', () => {
  state.filters = {
    sender:          $('filter-sender').value.trim(),
    recipient:       $('filter-recipient').value.trim(),
    subject:         $('filter-subject').value.trim(),
    date_from:       $('filter-date-from').value,
    date_to:         $('filter-date-to').value,
    has_attachments: $('filter-attachments').checked ? '1' : '',
  };
  state.page = 1;
  state.currentView = 'folder';
  loadEmailList();
});

$('filter-clear').addEventListener('click', () => {
  clearFilters();
  state.page = 1;
  loadEmailList();
});

function clearFilters() {
  ['filter-sender','filter-recipient','filter-subject',
   'filter-date-from','filter-date-to'].forEach(id => $(id).value = '');
  $('filter-attachments').checked = false;
  state.filters = {};
}

// ── Search ────────────────────────────────────────────────────────────────
function doSearch() {
  const q = $('search-input').value.trim();
  if (!q) return;
  state.searchQuery = q;
  state.currentView = 'search';
  state.page = 1;
  document.querySelectorAll('.folder-item').forEach(el => el.classList.remove('active'));
  loadEmailList();
}

$('search-btn').addEventListener('click', doSearch);
$('search-input').addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });

// ── Bookmarks ─────────────────────────────────────────────────────────────
document.querySelector('[data-view="bookmarks"]').addEventListener('click', () => {
  state.currentView = 'bookmarks';
  state.page = 1;
  document.querySelectorAll('.folder-item').forEach(el => el.classList.remove('active'));
  document.querySelector('[data-view="bookmarks"]').classList.add('active');
  loadEmailList();
});

async function renderBookmarks() {
  const container = $('email-list');
  const status = $('list-status');
  container.innerHTML = '';
  status.textContent = 'Loading…';
  try {
    const emails = await apiFetch('/api/bookmarks');
    status.textContent = '';
    if (emails.length === 0) { status.textContent = 'No bookmarks yet.'; return; }
    emails.forEach(e => {
      const li = document.createElement('li');
      li.className = 'email-item bookmarked';
      li.dataset.id = e.id;
      li.innerHTML = `
        <span class="ei-from">${esc(e.sender || '(unknown)')}</span>
        <span class="ei-date">${fmt_date(e.date_sent)}</span>
        <span class="ei-subject">${esc(e.subject || '(no subject)')}</span>
        <span class="ei-folder">${esc(e.folder_name || '')}</span>
      `;
      li.addEventListener('click', () => openEmail(e.id));
      container.appendChild(li);
    });
    updatePagination(emails.length);
  } catch (err) {
    status.textContent = 'Error: ' + err.message;
  }
}

// ── Open email ────────────────────────────────────────────────────────────
async function openEmail(emailId) {
  // Highlight in list
  document.querySelectorAll('.email-item').forEach(el => {
    el.classList.toggle('selected', el.dataset.id == emailId);
  });

  state.currentEmailId = emailId;
  $('view-placeholder').style.display = 'none';
  $('view-content').style.display = 'flex';

  try {
    const e = await apiFetch(`/api/email/${emailId}`);

    $('view-subject').textContent = e.subject || '(no subject)';
    $('view-from').textContent    = e.sender || '';
    $('view-to').textContent      = e.recipients || '';
    $('view-date').textContent    = e.date_sent ? new Date(e.date_sent).toLocaleString() : '';
    $('view-folder').textContent  = e.folder_name || '';

    // Bookmark button
    const btnBm = $('btn-bookmark');
    btnBm.classList.toggle('active', !!e.bookmarked);
    btnBm.title = e.bookmarked ? 'Remove bookmark' : 'Bookmark';

    // Download links
    $('btn-download-eml').onclick = () => window.open(`/api/email/${emailId}/download`);
    $('btn-download-txt').onclick = () => window.open(`/api/email/${emailId}/export/text`);

    // Tags
    renderTags(e.tags || []);

    // Attachments
    const attArea = $('attachments-area');
    const attList = $('attachments-list');
    if (e.attachments && e.attachments.length > 0) {
      attArea.style.display = '';
      attList.innerHTML = e.attachments.map(a => `
        <li class="att-item">
          📎 <a href="/api/attachment/${a.id}" download="${esc(a.filename)}">
            ${esc(a.filename || 'attachment')}
          </a>
          <span class="att-size">${fmt_size(a.size)}</span>
        </li>`).join('');
    } else {
      attArea.style.display = 'none';
    }

    // Body
    loadBody(emailId, state.bodyMode);

  } catch (err) {
    $('view-subject').textContent = 'Error loading email';
    console.error(err);
  }
}

function loadBody(emailId, mode) {
  state.bodyMode = mode;
  $('btn-view-html').classList.toggle('active', mode === 'html');
  $('btn-view-plain').classList.toggle('active', mode === 'plain');

  if (mode === 'html') {
    $('body-iframe').style.display = '';
    $('body-plain').style.display  = 'none';
    const highlight = state.currentView === 'search' ? state.searchQuery : '';
    const qs = highlight ? `?highlight=${encodeURIComponent(highlight)}` : '';
    $('body-iframe').src = `/api/email/${emailId}/html${qs}`;
  } else {
    $('body-iframe').style.display = 'none';
    $('body-plain').style.display  = '';
    apiFetch(`/api/email/${emailId}`).then(e => {
      $('body-plain').textContent = e.body_text || '(no plain text body)';
    });
  }
}

$('btn-view-html').addEventListener('click',  () => loadBody(state.currentEmailId, 'html'));
$('btn-view-plain').addEventListener('click', () => loadBody(state.currentEmailId, 'plain'));

// ── Bookmark toggle ───────────────────────────────────────────────────────
$('btn-bookmark').addEventListener('click', async () => {
  if (!state.currentEmailId) return;
  const isActive = $('btn-bookmark').classList.contains('active');
  const method = isActive ? 'DELETE' : 'POST';
  await fetch(`/api/email/${state.currentEmailId}/bookmark`, {method});
  $('btn-bookmark').classList.toggle('active', !isActive);
  // Refresh list to update star
  loadEmailList();
});

// ── Tags ──────────────────────────────────────────────────────────────────
function renderTags(tags) {
  const tagList = $('tag-list');
  tagList.innerHTML = tags.map(t =>
    `<span class="tag-chip">${esc(t)}
      <button class="remove-tag" data-tag="${esc(t)}" title="Remove tag">×</button>
    </span>`
  ).join('');
  tagList.querySelectorAll('.remove-tag').forEach(btn => {
    btn.addEventListener('click', () => removeTag(btn.dataset.tag));
  });
}

async function addTag() {
  const inp = $('tag-input');
  const tag = inp.value.trim();
  if (!tag || !state.currentEmailId) return;
  await fetch(`/api/email/${state.currentEmailId}/tags`, {
    method: 'POST', body: JSON.stringify({tag}), headers: {'Content-Type':'application/json'}
  });
  inp.value = '';
  const tags = await apiFetch(`/api/email/${state.currentEmailId}/tags`);
  renderTags(tags);
}

async function removeTag(tag) {
  if (!state.currentEmailId) return;
  await fetch(`/api/email/${state.currentEmailId}/tags`, {
    method: 'DELETE', body: JSON.stringify({tag}), headers: {'Content-Type':'application/json'}
  });
  const tags = await apiFetch(`/api/email/${state.currentEmailId}/tags`);
  renderTags(tags);
}

$('tag-add-btn').addEventListener('click', addTag);
$('tag-input').addEventListener('keydown', e => { if (e.key === 'Enter') addTag(); });

// ── Keyboard shortcuts ────────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  // Don't fire when focused in inputs
  if (['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName)) return;
  if (e.key === '/' || e.key === 's') { e.preventDefault(); $('search-input').focus(); }
});

// ── Boot ──────────────────────────────────────────────────────────────────
(async function init() {
  await Promise.all([loadStats(), loadFolders()]);
})();
