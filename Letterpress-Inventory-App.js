// ============================================================
// SPLASH SCREEN + PASSWORD PROTECTION
// ============================================================
(function() {
  const CORRECT_PASSWORD = 'iL0veHemr!4ever';
  const splash = document.getElementById('splash-screen');
  const stored = sessionStorage.getItem('pba_auth');
  const alreadyAuthed = stored === CORRECT_PASSWORD;

  function dismissSplash() {
    if (!splash) return;
    splash.classList.add('fade-out');
    setTimeout(() => splash.remove(), 500);
  }

  function promptPassword() {
    const input = prompt('🔐 Prints by Angel — Enter password:');
    if (input !== CORRECT_PASSWORD) {
      dismissSplash();
      document.body.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:Georgia,serif;color:#3D2B1F;background:#F5F0E4;font-size:1.1rem;">❌ Access denied.</div>';
      return false;
    }
    sessionStorage.setItem('pba_auth', CORRECT_PASSWORD);
    return true;
  }

  const MIN_SPLASH_MS = 3000;
  const splashStart = Date.now();

  function dismissAfterMinimum(callback) {
    const elapsed = Date.now() - splashStart;
    const remaining = Math.max(0, MIN_SPLASH_MS - elapsed);
    setTimeout(callback || dismissSplash, remaining);
  }

  if (alreadyAuthed) {
    // Already logged in — show splash for minimum time then dismiss
    dismissAfterMinimum();
  } else {
    // Show splash for minimum time, then prompt password
    dismissAfterMinimum(() => {
      if (!promptPassword()) return;
      dismissSplash();
    });
  }
})();
// ============================================================
// PRINTS BY ANGEL — Inventory App
// ============================================================

const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzM4FR2fzJ9LAE7w1R-J9u0-6V6HSBP5om2NsX59D8-fCH045rVvLvKMZBia8PX5zME/exec';
const appContainer = document.getElementById('app-container');

// ============================================================
// CACHE
// ============================================================
let itemsCache    = null;
let partnersCache = null;
let tagsCache     = null;
let machinesCache = null;

// ============================================================
// DOG MASCOT
// ============================================================
function dogLoading(message = 'Loading...') {
  return `
    <div class="dog-state">
      <div class="dog-png-wrap">
        <img src="doggycutetonguev2.png" class="dog-png dog-bounce" alt="Loading..." onerror="this.style.display='none';this.parentNode.innerHTML='🐾'">
      </div>
      <div class="dog-dots">
        <span class="dot"></span><span class="dot"></span><span class="dot"></span>
      </div>
      <div class="dog-message">${message}</div>
    </div>`;
}

function dogError(message = 'Something went wrong.') {
  return `
    <div class="dog-state">
      <div class="dog-png-wrap" style="position:relative;">
        <img src="doggycutetonguev2.png" class="dog-png" alt="Error" style="filter:grayscale(30%);" onerror="this.style.display='none';this.parentNode.innerHTML='🐾'">
        <div class="dog-badge dog-badge-error">WTF?!<br><span>can't connect</span></div>
      </div>
      <div class="dog-message error">${message}</div>
    </div>`;
}

function dogEmpty(message = 'Nothing here yet.') {
  return `
    <div class="dog-state">
      <div class="dog-png-wrap" style="position:relative;">
        <img src="doggycutetonguev2.png" class="dog-png" alt="Empty" onerror="this.style.display='none';this.parentNode.innerHTML='🐾'">
        <div class="dog-badge dog-badge-empty">
          <svg width="38" height="38" viewBox="0 0 38 38" fill="none">
            <circle cx="16" cy="16" r="11" stroke="#4AABAB" stroke-width="3.5" fill="rgba(74,171,171,0.1)"/>
            <line x1="24" y1="24" x2="34" y2="34" stroke="#4AABAB" stroke-width="3.5" stroke-linecap="round"/>
            <text x="10" y="21" font-size="11" fill="#4AABAB" font-weight="bold">?</text>
          </svg>
        </div>
      </div>
      <div class="dog-message">${message}</div>
    </div>`;
}

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  setupNavigation();
  setupMobileMenu();
  setupDetailPanel();
  setupModal();
  const lastPage = sessionStorage.getItem('pba_current_page') || 'home';
  loadPage(lastPage);
  document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
  document.querySelector(`[data-page="${lastPage}"]`)?.classList.add('active');
});

// ============================================================
// NAVIGATION
// ============================================================
function setupNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(btn => {
    btn.addEventListener('click', () => {
      const page = btn.dataset.page;
      navItems.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      loadPage(page);
      closeMobileMenu();
    });
  });
}

function setupMobileMenu() {
  const hamburger = document.getElementById('hamburger-btn');
  const sidebar   = document.getElementById('sidebar');
  const overlay   = document.getElementById('sidebar-overlay');
  hamburger?.addEventListener('click', () => {
    const isOpen = sidebar.classList.contains('open');
    if (isOpen) { closeMobileMenu(); }
    else {
      sidebar.classList.add('open');
      overlay.classList.add('open');
      hamburger.classList.add('open');
    }
  });
  overlay?.addEventListener('click', closeMobileMenu);
}

function closeMobileMenu() {
  document.getElementById('sidebar')?.classList.remove('open');
  document.getElementById('sidebar-overlay')?.classList.remove('open');
  document.getElementById('hamburger-btn')?.classList.remove('open');
}

// ============================================================
// DETAIL PANEL
// ============================================================
function setupDetailPanel() {
  document.getElementById('detail-close')?.addEventListener('click', closeDetailPanel);
  document.getElementById('detail-overlay')?.addEventListener('click', closeDetailPanel);
}

function openDetailPanel(html) {
  document.getElementById('detail-panel-content').innerHTML = html;
  document.getElementById('detail-panel').classList.add('open');
  document.getElementById('detail-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeDetailPanel() {
  document.getElementById('detail-panel').classList.remove('open');
  document.getElementById('detail-overlay').classList.remove('open');
  document.body.style.overflow = '';
}

// ============================================================
// MODAL
// ============================================================
function setupModal() {
  document.getElementById('modal-close')?.addEventListener('click', closeModal);
  document.getElementById('modal-overlay')?.addEventListener('click', (e) => {
    if (e.target === document.getElementById('modal-overlay')) closeModal();
  });
}

function openModal(html) {
  document.getElementById('modal-content').innerHTML = html;
  document.getElementById('modal-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('open');
  document.body.style.overflow = '';
}

function openSubModal(html) {
  document.getElementById('sub-modal-content').innerHTML = html;
  document.getElementById('sub-modal-overlay').style.display = '';
  document.getElementById('sub-modal-overlay').classList.add('open');
}

function closeSubModal() {
  document.getElementById('sub-modal-overlay').classList.remove('open');
  document.getElementById('sub-modal-overlay').style.display = 'none';
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('sub-modal-close')?.addEventListener('click', closeSubModal);
  document.getElementById('sub-modal-overlay')?.addEventListener('click', (e) => {
    if (e.target === document.getElementById('sub-modal-overlay')) closeSubModal();
  });
});

// ============================================================
// TOAST
// ============================================================
// Fix old Google Drive photo URLs that no longer work as image sources
function fixPhotoUrl(url) {
  if (!url) return '';
  // Convert drive.google.com/uc?export=view&id=XXX to lh3 format
  const match1 = url.match(/drive\.google\.com\/uc\?.*id=([a-zA-Z0-9_-]+)/);
  if (match1) return 'https://lh3.googleusercontent.com/d/' + match1[1];
  // Convert drive.google.com/file/d/XXX/view to lh3 format
  const match2 = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match2) return 'https://lh3.googleusercontent.com/d/' + match2[1];
  // Convert drive.google.com/open?id=XXX to lh3 format
  const match3 = url.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (match3) return 'https://lh3.googleusercontent.com/d/' + match3[1];
  return url;
}

// Live-convert pasted Google Drive URLs and show preview
function handlePhotoUrlPaste(inputEl, previewSelector) {
  const raw = inputEl.value.trim();
  const fixed = fixPhotoUrl(raw);
  if (fixed !== raw && fixed !== '') {
    inputEl.value = fixed;
  }
  // Show/update preview
  let previewWrap = document.querySelector(previewSelector);
  if (fixed) {
    if (!previewWrap) {
      previewWrap = document.createElement('div');
      previewWrap.className = previewSelector.replace('.', '') + ' photo-preview-wrap';
      inputEl.before(previewWrap);
    }
    previewWrap.innerHTML = `<img class="photo-preview-thumb" src="${fixed}" alt="Photo preview" onerror="this.parentElement.innerHTML='<span style=\\'color:var(--amber);font-size:0.75rem\\'>⚠️ Could not load image</span>'">`;
  } else if (previewWrap) {
    previewWrap.innerHTML = '';
  }
}

function formatPhoneField(input) {
  const digits = input.value.replace(/\D/g, '').slice(0, 10);
  if (digits.length <= 3) input.value = digits;
  else if (digits.length <= 6) input.value = digits.slice(0,3) + '-' + digits.slice(3);
  else input.value = digits.slice(0,3) + '-' + digits.slice(3,6) + '-' + digits.slice(6);
}

function addSkuCodeRow(product = '', code = '') {
  const list = document.getElementById('sku-codes-list');
  if (!list) return;
  const row = document.createElement('div');
  row.className = 'sku-code-row';
  row.style.cssText = 'display:flex;gap:0.4rem;align-items:center;margin-bottom:0.35rem;';
  row.innerHTML = `
    <input class="field-input sku-product" type="text" placeholder="Product (e.g. VA Shop Cards)" value="${product}" style="flex:1;">
    <input class="field-input sku-code" type="text" placeholder="Code (e.g. 8440)" value="${code}" style="width:5rem;">
    <button type="button" style="background:none;border:none;color:var(--coral);font-size:1.2rem;cursor:pointer;padding:0 0.3rem;" onclick="this.parentElement.remove()">✕</button>`;
  list.appendChild(row);
}

function showToast(message, type = '') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => { toast.className = 'toast'; }, 3000);
}

// ============================================================
// PAGE ROUTER
// ============================================================
function loadPage(pageName) {
  sessionStorage.setItem('pba_current_page', pageName);
  const template = document.getElementById('record-sale-page');
  if (template) template.style.display = 'none';
  appContainer.innerHTML = '';
  closeDetailPanel();
  closeModal();

  switch (pageName) {
    case 'home':                renderHomePage(); break;
    case 'master-items':        renderMasterItemsPage(); break;
    case 'new-item-design':     renderNewItemDesignPage(); break;
    case 'print-stock-updater': renderPrintStockUpdaterPage(); break;
    case 'record-sale':         renderRecordSalePage(); break;
    case 'retail-partners':     renderRetailPartnersPage(); break;
    case 'vending-machines':    renderVendingMachinesPage(); break;
    case 'orders':              renderOrdersPage(); break;
    case 'market-sales':        renderMarketSalesPage(); break;
    case 'sales-reports':       renderSalesReportsPage(); break;
    case 'inventory-auditor':   renderInventoryAuditorPage(); break;
    case 'settings':            renderSettingsPage(); break;
    default:
      appContainer.innerHTML = `<div class="dog-state" style="padding:4rem">${dogEmpty('Page not found')}</div>`;
  }
}

// ============================================================
// DASHBOARD — CONSTANTS & PREFS
// ============================================================
const DASH_PREFS_KEY = 'pba_dash_prefs_v3';

const ALL_STAT_TILES = [
  { id: 'stat-revenue',  icon: '💰', color: 'teal',  label: 'Total Revenue'   },
  { id: 'stat-sold',     icon: '🃏', color: 'amber', label: 'Cards Sold'      },
  { id: 'stat-stock',    icon: '📦', color: 'green', label: 'In Stock (Home)' },
  { id: 'stat-consign',  icon: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="28" height="28"><rect x="10" y="30" width="44" height="26" rx="2" fill="#f5e6a3"/><rect x="10" y="30" width="44" height="26" rx="2" fill="none" stroke="#2d2d2d" stroke-width="2.5"/><rect x="24" y="38" width="12" height="18" rx="1.5" fill="#f5e6a3" stroke="#2d2d2d" stroke-width="2"/><path d="M36 47 a1.2 1.2 0 1 1 0 0.1" fill="#2d2d2d"/><rect x="40" y="38" width="9" height="8" rx="1" fill="#f5e6a3" stroke="#2d2d2d" stroke-width="1.5"/><path d="M7 30 L32 20 L57 30" fill="#6b88b0" stroke="#2d2d2d" stroke-width="2.5" stroke-linejoin="round"/><line x1="14" y1="28" x2="14" y2="22" stroke="white" stroke-width="3.5" opacity="0.5"/><line x1="22" y1="25.5" x2="22" y2="20.5" stroke="white" stroke-width="3.5" opacity="0.5"/><line x1="42" y1="25.5" x2="42" y2="20.5" stroke="white" stroke-width="3.5" opacity="0.5"/><line x1="50" y1="28" x2="50" y2="22" stroke="white" stroke-width="3.5" opacity="0.5"/><path d="M7 30 Q10.5 34 14 30 Q17.5 34 21 30 Q24.5 34 28 30 Q31.5 34 35 30 Q38.5 34 42 30 Q45.5 34 49 30 Q52.5 34 57 30" fill="none" stroke="#2d2d2d" stroke-width="2"/><rect x="22" y="12" width="20" height="9" rx="2" fill="white" stroke="#2d2d2d" stroke-width="2"/></svg>`, color: 'coral', label: 'On Consignment'  },
  { id: 'stat-designs',  icon: '✏️', color: 'teal',  label: 'Total Designs'   },
  { id: 'stat-printed',  icon: '🖨️', color: 'amber', label: 'Total Printed'   },
  { id: 'stat-partners', icon: '🤝', color: 'green', label: 'Retail Partners' },
  { id: 'stat-machines', icon: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="28" height="28"><rect x="12" y="6" width="40" height="52" rx="3" fill="#e03030" stroke="#2d2d2d" stroke-width="2"/><rect x="16" y="10" width="24" height="34" rx="2" fill="#f5f0e8"/><rect x="16" y="10" width="24" height="34" rx="2" fill="none" stroke="#2d2d2d" stroke-width="1.5"/><line x1="16" y1="21" x2="40" y2="21" stroke="#ddd" stroke-width="1"/><line x1="16" y1="32" x2="40" y2="32" stroke="#ddd" stroke-width="1"/><circle cx="21" cy="16" r="2.5" fill="#5bc0de"/><circle cx="28" cy="16" r="2.5" fill="#5bc0de"/><circle cx="35" cy="16" r="2.5" fill="#f0c040"/><circle cx="21" cy="27" r="2.5" fill="#e05050"/><circle cx="28" cy="27" r="2.5" fill="#50b080"/><circle cx="35" cy="27" r="2.5" fill="#e05050"/><circle cx="21" cy="38" r="2.5" fill="#f0c040"/><circle cx="28" cy="38" r="2.5" fill="#5bc0de"/><circle cx="35" cy="38" r="2.5" fill="#50b080"/><rect x="42" y="12" width="7" height="16" rx="1.5" fill="#1a1a1a"/><rect x="16" y="48" width="24" height="6" rx="1" fill="#b02020"/><rect x="16" y="48" width="24" height="6" rx="1" fill="none" stroke="#2d2d2d" stroke-width="1"/></svg>`, color: 'coral', label: 'Active Machines' },
];

const ALL_QUICK_ACTIONS = [
  { id: 'qa-add-design',    icon: '✚',  label: 'Add New Design',   page: 'new-item-design'     },
  { id: 'qa-print-run',     icon: '🖨️', label: 'Log Print Run',    page: 'print-stock-updater' },
  { id: 'qa-retail-update', icon: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="28" height="28"><rect x="10" y="30" width="44" height="26" rx="2" fill="#f5e6a3"/><rect x="10" y="30" width="44" height="26" rx="2" fill="none" stroke="#2d2d2d" stroke-width="2.5"/><rect x="24" y="38" width="12" height="18" rx="1.5" fill="#f5e6a3" stroke="#2d2d2d" stroke-width="2"/><path d="M36 47 a1.2 1.2 0 1 1 0 0.1" fill="#2d2d2d"/><rect x="40" y="38" width="9" height="8" rx="1" fill="#f5e6a3" stroke="#2d2d2d" stroke-width="1.5"/><path d="M7 30 L32 20 L57 30" fill="#6b88b0" stroke="#2d2d2d" stroke-width="2.5" stroke-linejoin="round"/><line x1="14" y1="28" x2="14" y2="22" stroke="white" stroke-width="3.5" opacity="0.5"/><line x1="22" y1="25.5" x2="22" y2="20.5" stroke="white" stroke-width="3.5" opacity="0.5"/><line x1="42" y1="25.5" x2="42" y2="20.5" stroke="white" stroke-width="3.5" opacity="0.5"/><line x1="50" y1="28" x2="50" y2="22" stroke="white" stroke-width="3.5" opacity="0.5"/><path d="M7 30 Q10.5 34 14 30 Q17.5 34 21 30 Q24.5 34 28 30 Q31.5 34 35 30 Q38.5 34 42 30 Q45.5 34 49 30 Q52.5 34 57 30" fill="none" stroke="#2d2d2d" stroke-width="2"/><rect x="22" y="12" width="20" height="9" rx="2" fill="white" stroke="#2d2d2d" stroke-width="2"/></svg>`, label: 'Retail Stock Update', page: 'record-sale'         },
  { id: 'qa-orders',        icon: '📋', label: 'View Orders',      page: 'orders'              },
  { id: 'qa-vending',       icon: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="28" height="28"><rect x="12" y="6" width="40" height="52" rx="3" fill="#e03030" stroke="#2d2d2d" stroke-width="2"/><rect x="16" y="10" width="24" height="34" rx="2" fill="#f5f0e8"/><rect x="16" y="10" width="24" height="34" rx="2" fill="none" stroke="#2d2d2d" stroke-width="1.5"/><line x1="16" y1="21" x2="40" y2="21" stroke="#ddd" stroke-width="1"/><line x1="16" y1="32" x2="40" y2="32" stroke="#ddd" stroke-width="1"/><circle cx="21" cy="16" r="2.5" fill="#5bc0de"/><circle cx="28" cy="16" r="2.5" fill="#5bc0de"/><circle cx="35" cy="16" r="2.5" fill="#f0c040"/><circle cx="21" cy="27" r="2.5" fill="#e05050"/><circle cx="28" cy="27" r="2.5" fill="#50b080"/><circle cx="35" cy="27" r="2.5" fill="#e05050"/><circle cx="21" cy="38" r="2.5" fill="#f0c040"/><circle cx="28" cy="38" r="2.5" fill="#5bc0de"/><circle cx="35" cy="38" r="2.5" fill="#50b080"/><rect x="42" y="12" width="7" height="16" rx="1.5" fill="#1a1a1a"/><rect x="16" y="48" width="24" height="6" rx="1" fill="#b02020"/><rect x="16" y="48" width="24" height="6" rx="1" fill="none" stroke="#2d2d2d" stroke-width="1"/></svg>`, label: 'Vending Machines', page: 'vending-machines'    },
  { id: 'qa-partners',      icon: '❖',  label: 'Partners',         page: 'retail-partners'     },
  { id: 'qa-market-sale',   icon: '🎪', label: 'Log Market Sale',  page: 'market-sales'        },
];

function getDashPrefs() {
  try { return JSON.parse(localStorage.getItem(DASH_PREFS_KEY) || 'null') || getDefaultDashPrefs(); }
  catch (e) { return getDefaultDashPrefs(); }
}
function getDefaultDashPrefs() {
  return {
    statOrder:  ALL_STAT_TILES.map(t => t.id),
    statHidden: [],
    qaOrder:    ALL_QUICK_ACTIONS.map(a => a.id),
    qaHidden:   [],
  };
}
function saveDashPrefs(prefs) {
  localStorage.setItem(DASH_PREFS_KEY, JSON.stringify(prefs));
}

// ============================================================
// DASHBOARD — RENDER
// ============================================================
async function renderHomePage() {
  const prefs       = getDashPrefs();
  const orderedStats = prefs.statOrder.map(id => ALL_STAT_TILES.find(t => t.id === id)).filter(Boolean);
  const orderedQA    = prefs.qaOrder.map(id => ALL_QUICK_ACTIONS.find(a => a.id === id)).filter(Boolean);

  appContainer.innerHTML = `
    <div class="page-header">
      <div>
        <h1 class="page-title">Dashboard</h1>
        <p class="page-subtitle">Welcome back — here's what's happening with Prints by Angel</p>
      </div>
      <div class="page-actions">
        <button class="btn btn-secondary btn-sm" onclick="openDashCustomizePanel()">⊹ Customize</button>
        <button class="btn btn-secondary btn-sm" onclick="openLinksModal()">🔗 Links</button>
      </div>
    </div>

    <h3 class="section-title" style="margin-bottom:0.5rem;">At a Glance</h3>
    <div class="stats-grid" id="stats-grid">
      ${orderedStats.map(s => `
        <div class="stat-card ${prefs.statHidden.includes(s.id) ? 'tile-hidden' : ''}" id="tile-${s.id}">
          <div class="stat-icon ${s.color}">${s.icon}</div>
          <div class="stat-value" id="${s.id}">...</div>
          <div class="stat-label">${s.label}</div>
        </div>`).join('')}
    </div>

    <h3 class="section-title" style="margin:1.5rem 0 0.5rem;">Quick Actions</h3>
    <div class="quick-actions" id="quick-actions-grid">
      ${orderedQA.map(a => `
        <button class="quick-action-btn ${prefs.qaHidden.includes(a.id) ? 'tile-hidden' : ''}"
                id="${a.id}"
                onclick="document.querySelector('[data-page=${a.page}]').click()">
          <span class="quick-action-icon">${a.icon}</span>
          <span class="quick-action-label">${a.label}</span>
        </button>`).join('')}
    </div>

    <!-- Customize side panel -->
    <div class="customize-overlay" id="customize-overlay" onclick="closeDashCustomizePanel()"></div>
    <div class="customize-panel" id="customize-panel">
      <div class="customize-panel-header">
        <span>⊹ Customize Dashboard</span>
        <button onclick="closeDashCustomizePanel()">✕</button>
      </div>
      <div class="customize-panel-body" id="customize-panel-body"></div>
    </div>`;

  await loadDashboardStats();
}

// ============================================================
// DASHBOARD — CUSTOMIZE PANEL
// ============================================================
function openDashCustomizePanel() {
  renderCustomizePanelBody();
  document.getElementById('customize-panel').classList.add('open');
  document.getElementById('customize-overlay').classList.add('open');
}

function closeDashCustomizePanel() {
  document.getElementById('customize-panel')?.classList.remove('open');
  document.getElementById('customize-overlay')?.classList.remove('open');
}

function renderCustomizePanelBody() {
  const prefs = getDashPrefs();
  const body  = document.getElementById('customize-panel-body');
  if (!body) return;

  const makeList = (items, section, hiddenList) => items.map(item => {
    const hidden = hiddenList.includes(item.id);
    return `
      <div class="sortable-item ${hidden ? 'sortable-hidden' : ''}"
           data-id="${item.id}" data-section="${section}"
           draggable="true"
           ondragstart="onDragStart(event)"
           ondragover="onDragOver(event)"
           ondrop="onDrop(event)"
           ondragend="onDragEnd(event)">
        <span class="drag-handle">⠿</span>
        <span class="sortable-icon">${item.icon}</span>
        <span class="sortable-label">${item.label}</span>
        <button class="sortable-toggle-btn"
                onclick="toggleTileVisibility('${item.id}','${section}',this)">
          ${hidden ? '👁' : '✕'}
        </button>
      </div>`;
  }).join('');

  const orderedStats = prefs.statOrder.map(id => ALL_STAT_TILES.find(t => t.id === id)).filter(Boolean);
  const orderedQA    = prefs.qaOrder.map(id => ALL_QUICK_ACTIONS.find(a => a.id === id)).filter(Boolean);

  body.innerHTML = `
    <div class="customize-section-title">📊 At a Glance Tiles</div>
    <p class="customize-hint">Drag to reorder · ✕ hides · 👁 restores</p>
    <div class="sortable-list">
      ${makeList(orderedStats, 'stats', prefs.statHidden)}
    </div>

    <div class="customize-divider"></div>

    <div class="customize-section-title">⚡ Quick Actions</div>
    <p class="customize-hint">Drag to reorder · ✕ hides · 👁 restores</p>
    <div class="sortable-list">
      ${makeList(orderedQA, 'qa', prefs.qaHidden)}
    </div>

    <div class="customize-footer">
      <button class="btn btn-primary" style="width:100%" onclick="saveDashLayout()">✓ Save Layout</button>
      <button class="btn btn-secondary" style="width:100%;margin-top:0.5rem" onclick="resetDashLayout()">↺ Reset to Defaults</button>
    </div>`;
}

// ============================================================
// DASHBOARD — DRAG & DROP
// ============================================================
let _dragId      = null;
let _dragSection = null;

function onDragStart(e) {
  _dragId      = e.currentTarget.dataset.id;
  _dragSection = e.currentTarget.dataset.section;
  e.currentTarget.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
}
function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  document.querySelectorAll('.sortable-item.drag-over').forEach(el => el.classList.remove('drag-over'));
  e.currentTarget.classList.add('drag-over');
}
function onDrop(e) {
  e.preventDefault();
  const targetId      = e.currentTarget.dataset.id;
  const targetSection = e.currentTarget.dataset.section;
  if (!targetId || targetId === _dragId || targetSection !== _dragSection) return;

  const prefs    = getDashPrefs();
  const orderKey = _dragSection === 'stats' ? 'statOrder' : 'qaOrder';
  const arr      = [...prefs[orderKey]];
  const from     = arr.indexOf(_dragId);
  const to       = arr.indexOf(targetId);
  if (from === -1 || to === -1) return;
  arr.splice(from, 1);
  arr.splice(to, 0, _dragId);
  prefs[orderKey] = arr;
  saveDashPrefs(prefs);
  renderCustomizePanelBody();
  applyDashLayout();
}
function onDragEnd() {
  document.querySelectorAll('.sortable-item').forEach(el => el.classList.remove('dragging', 'drag-over'));
  _dragId = null;
}

// ============================================================
// DASHBOARD — TOGGLE VISIBILITY
// ============================================================
function toggleTileVisibility(id, section, btn) {
  const prefs     = getDashPrefs();
  const hiddenKey = section === 'stats' ? 'statHidden' : 'qaHidden';
  const idx       = prefs[hiddenKey].indexOf(id);
  if (idx === -1) { prefs[hiddenKey].push(id); }
  else            { prefs[hiddenKey].splice(idx, 1); }
  saveDashPrefs(prefs);
  renderCustomizePanelBody();
  applyDashLayout();
}

// ============================================================
// DASHBOARD — SAVE / RESET
// ============================================================
function saveDashLayout() {
  closeDashCustomizePanel();
  showToast('Dashboard saved!', 'success');
}
function resetDashLayout() {
  saveDashPrefs(getDefaultDashPrefs());
  closeDashCustomizePanel();
  renderHomePage();
  showToast('Reset to defaults', '');
}

function applyDashLayout() {
  const prefs     = getDashPrefs();
  const statsGrid = document.getElementById('stats-grid');
  const qaGrid    = document.getElementById('quick-actions-grid');
  if (!statsGrid || !qaGrid) return;

  prefs.statOrder.forEach(id => {
    const el = document.getElementById('tile-' + id);
    if (el) statsGrid.appendChild(el);
  });
  ALL_STAT_TILES.forEach(t => {
    const el = document.getElementById('tile-' + t.id);
    if (el) el.classList.toggle('tile-hidden', prefs.statHidden.includes(t.id));
  });

  prefs.qaOrder.forEach(id => {
    const el = document.getElementById(id);
    if (el) qaGrid.appendChild(el);
  });
  ALL_QUICK_ACTIONS.forEach(a => {
    const el = document.getElementById(a.id);
    if (el) el.classList.toggle('tile-hidden', prefs.qaHidden.includes(a.id));
  });
}

// ============================================================
// DASHBOARD — LINKS MODAL
// ============================================================
const RETAILER_PAGE_URL   = 'https://retailers.printsbyangel.com/';
const PARTNER_REQUEST_URL = 'https://forms.gle/PQJk2YXAdvPbr53a6';

function openLinksModal() {
  openModal(`
    <div class="modal-title">🔗 Share Links</div>
    <p style="font-size:0.875rem;color:var(--brown-mid);margin-bottom:1.25rem;">
      Share these with your partners or post them anywhere.
    </p>
    <div class="link-card">
      <div class="link-card-label">🛒 Retailer Ordering Page</div>
      <div class="link-card-url" id="retailer-url">${RETAILER_PAGE_URL}</div>
      <div class="link-card-actions">
        <button class="btn btn-primary btn-sm" onclick="copyLink('retailer-url')">📋 Copy</button>
        <button class="btn btn-secondary btn-sm" onclick="window.open('${RETAILER_PAGE_URL}','_blank')">↗ Open</button>
      </div>
      <div class="link-qr-wrap" id="qr-retailer"></div>
    </div>
    <div class="link-card" style="margin-top:1rem;">
      <div class="link-card-label">📝 Partner Account Request Form</div>
      <div class="link-card-url" id="partner-req-url">${PARTNER_REQUEST_URL}</div>
      <div class="link-card-actions">
        <button class="btn btn-primary btn-sm" onclick="copyLink('partner-req-url')">📋 Copy</button>
        <button class="btn btn-secondary btn-sm" onclick="window.open('${PARTNER_REQUEST_URL}','_blank')">↗ Open</button>
      </div>
      <div class="link-qr-wrap" id="qr-partner"></div>
    </div>
    <div id="link-copy-toast" style="display:none;margin-top:0.75rem;text-align:center;font-size:0.875rem;color:var(--teal);font-weight:600;">✅ Copied!</div>`);

  generateQR('qr-retailer', RETAILER_PAGE_URL);
  generateQR('qr-partner',  PARTNER_REQUEST_URL);
}

function copyLink(elementId) {
  const url = document.getElementById(elementId)?.textContent?.trim();
  if (!url) return;
  navigator.clipboard.writeText(url).then(() => {
    const t = document.getElementById('link-copy-toast');
    if (t) { t.style.display = 'block'; setTimeout(() => t.style.display = 'none', 2000); }
  }).catch(() => {
    const el = document.createElement('textarea');
    el.value = url; document.body.appendChild(el); el.select();
    document.execCommand('copy'); document.body.removeChild(el);
    showToast('Copied!', 'success');
  });
}

function generateQR(containerId, url) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const img = document.createElement('img');
  img.src = `https://api.qrserver.com/v1/create-qr-code/?size=140x140&data=${encodeURIComponent(url)}&bgcolor=F5F0E4&color=2C1F17`;
  img.alt = 'QR Code';
  img.style.cssText = 'width:140px;height:140px;border-radius:8px;margin-top:0.75rem;display:block;';
  container.appendChild(img);
}

// ============================================================
// DASHBOARD — LOAD STATS
// ============================================================
async function loadDashboardStats() {
  try {
    const r    = await fetch(`${GOOGLE_SCRIPT_URL}?action=getDashboardStats`);
    const data = await r.json();
    if (data.success) {
      const s   = data.stats;
      const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
      set('stat-revenue',  `$${Number(s.totalRevenue || 0).toFixed(2)}`);
      set('stat-sold',     s.totalSold           || 0);
      set('stat-stock',    s.totalInStock         || 0);
      set('stat-consign',  s.totalOnConsignment   || 0);
      set('stat-designs',  s.totalDesigns         || 0);
      set('stat-printed',  s.totalPrinted         || 0);
      set('stat-partners', s.totalPartners        || 0);
      set('stat-machines', s.totalMachines        || 0);
      applyDashLayout();
      return;
    }
  } catch (e) {}

  ['stat-revenue','stat-sold','stat-stock','stat-consign','stat-printed','stat-machines'].forEach(id => {
    const el = document.getElementById(id);
    if (el && el.textContent === '...') el.textContent = '—';
  });
  try {
    if (!itemsCache) {
      const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getItems`);
      const d = await r.json();
      if (d.success) itemsCache = d.items.filter(i => i.ItemID);
    }
    const el = document.getElementById('stat-designs');
    if (el) el.textContent = itemsCache?.length || 0;
  } catch (e) {}
  try {
    const r2 = await fetch(`${GOOGLE_SCRIPT_URL}?action=getRetailPartners`);
    const d2 = await r2.json();
    const el  = document.getElementById('stat-partners');
    if (el && d2.success) el.textContent = d2.partners?.length || 0;
  } catch (e) {}
  applyDashLayout();
}

// ============================================================
// MAIN INVENTORY LIST
// ============================================================
async function renderMasterItemsPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div>
        <h1 class="page-title">Main Inventory List</h1>
        <p class="page-subtitle">All designs in your collection</p>
      </div>
      <div class="page-actions">
        <button class="btn btn-secondary" onclick="itemsCache=null;renderMasterItemsPage()" title="Refresh list">↺ Refresh</button>
        <button class="btn btn-primary" onclick="window._cameFromInventory=true;document.querySelector('[data-page=new-item-design]').click()">✚ Add Design</button>
      </div>
    </div>
    <div class="items-toolbar">
      <div class="search-box">
        <span class="search-icon">⌕</span>
        <input type="text" id="items-search" placeholder="Search by name or number...">
      </div>
      <button class="filter-btn" id="filter-btn" onclick="openFilterPanel()">
        <svg width="15" height="15" viewBox="0 0 15 15" fill="none"><path d="M1 3.5h13M3.5 7.5h8M6 11.5h3" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
        Filter
        <span class="filter-count-badge" id="filter-count" style="display:none">0</span>
      </button>
      <div class="view-toggle">
        <button class="view-toggle-btn active" id="grid-view-btn" title="Grid view">⊞</button>
        <button class="view-toggle-btn" id="list-view-btn" title="List view">☰</button>
      </div>
    </div>
    <div id="active-filters-bar"></div>
    <div id="items-container">${dogLoading('Loading designs...')}</div>

    <div class="filter-panel-overlay" id="filter-panel-overlay" onclick="closeFilterPanel()"></div>
    <div class="filter-panel" id="filter-panel">
      <div class="filter-panel-header">
        <span class="filter-panel-title">Filter Designs</span>
        <button class="filter-panel-close" onclick="closeFilterPanel()">✕</button>
      </div>
      <div class="filter-panel-body" id="filter-panel-body">Loading...</div>
      <div class="filter-panel-footer">
        <button class="btn btn-secondary" onclick="clearAllFilters()">Clear All</button>
        <button class="btn btn-primary" onclick="applyFilters()">Show Results</button>
      </div>
    </div>`;

  document.getElementById('grid-view-btn').addEventListener('click', () => {
    document.getElementById('grid-view-btn').classList.add('active');
    document.getElementById('list-view-btn').classList.remove('active');
    renderItemsGrid(currentItems);
  });
  document.getElementById('list-view-btn').addEventListener('click', () => {
    document.getElementById('list-view-btn').classList.add('active');
    document.getElementById('grid-view-btn').classList.remove('active');
    renderItemsList(currentItems);
  });

  await fetchAndDisplayItems();
  loadFilterPanel();

  document.getElementById('items-search')?.addEventListener('input', (e) => {
    const q        = e.target.value.toLowerCase();
    const filtered = (itemsCache || []).filter(item =>
      (item.DisplayName || item.Name || '').toLowerCase().includes(q) ||
      String(item.ItemID).includes(q)
    );
    currentItems = filtered;
    const isGrid = document.getElementById('grid-view-btn').classList.contains('active');
    isGrid ? renderItemsGrid(filtered) : renderItemsList(filtered);
  });
}

let currentItems    = [];
let activeTagFilters = new Set();

async function loadFilterPanel() {
  if (!tagsCache) {
    try {
      const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getTags`);
      const d = await r.json();
      if (d.success) tagsCache = d.tags;
    } catch (e) { return; }
  }
  renderFilterPanelBody();
}

function renderFilterPanelBody() {
  const body = document.getElementById('filter-panel-body');
  if (!body || !tagsCache) return;
  const categories = {};
  (tagsCache || []).forEach(tag => {
    if (!categories[tag.Category]) categories[tag.Category] = [];
    categories[tag.Category].push(tag);
  });
  body.innerHTML = Object.entries(categories).map(([cat, tags]) => `
    <div class="filter-category">
      <div class="filter-category-label">${cat}</div>
      <div class="filter-tag-chips">
        ${tags.map(tag => `
          <button class="filter-tag-chip ${activeTagFilters.has(tag.TagID) ? 'selected' : ''}"
            onclick="toggleTagFilter('${tag.TagID}', this)">
            ${tag.TagName}
          </button>`).join('')}
      </div>
    </div>`).join('');
}

function toggleTagFilter(tagId, btn) {
  if (activeTagFilters.has(tagId)) { activeTagFilters.delete(tagId); btn.classList.remove('selected'); }
  else { activeTagFilters.add(tagId); btn.classList.add('selected'); }
}

function openFilterPanel() {
  renderFilterPanelBody();
  document.getElementById('filter-panel').classList.add('open');
  document.getElementById('filter-panel-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeFilterPanel() {
  document.getElementById('filter-panel').classList.remove('open');
  document.getElementById('filter-panel-overlay').classList.remove('open');
  document.body.style.overflow = '';
}

function clearAllFilters() {
  activeTagFilters.clear();
  renderFilterPanelBody();
  updateActiveFiltersBar();
  currentItems = itemsCache || [];
  const isGrid = document.getElementById('grid-view-btn')?.classList.contains('active');
  isGrid ? renderItemsGrid(currentItems) : renderItemsList(currentItems);
}

function applyFilters() {
  closeFilterPanel();
  updateActiveFiltersBar();
  currentItems = itemsCache || [];
  if (activeTagFilters.size > 0) {
    currentItems = currentItems.filter(item =>
      (item._tagIds || []).some(id => activeTagFilters.has(id))
    );
  }
  const count = document.getElementById('filter-count');
  if (count) {
    if (activeTagFilters.size > 0) {
      count.textContent = activeTagFilters.size;
      count.style.display = 'inline-flex';
      document.getElementById('filter-btn')?.classList.add('has-filters');
    } else {
      count.style.display = 'none';
      document.getElementById('filter-btn')?.classList.remove('has-filters');
    }
  }
  const isGrid = document.getElementById('grid-view-btn')?.classList.contains('active');
  isGrid ? renderItemsGrid(currentItems) : renderItemsList(currentItems);
}

function updateActiveFiltersBar() {
  const bar = document.getElementById('active-filters-bar');
  if (!bar) return;
  if (activeTagFilters.size === 0) { bar.innerHTML = ''; return; }
  const tagNames = [...activeTagFilters].map(id => {
    const tag = (tagsCache || []).find(t => t.TagID === id);
    return tag ? `<span class="active-filter-chip">${tag.TagName} <button onclick="removeTagFilter('${id}')">✕</button></span>` : '';
  }).join('');
  bar.innerHTML = `<div class="active-filters-row">${tagNames}</div>`;
}

function removeTagFilter(tagId) {
  activeTagFilters.delete(tagId);
  applyFilters();
  renderFilterPanelBody();
}

async function fetchAndDisplayItems() {
  try {
    if (!itemsCache) {
      const response = await fetch(`${GOOGLE_SCRIPT_URL}?action=getItems`);
      const data     = await response.json();
      if (data.success) itemsCache = data.items.filter(i => i.ItemID);
    }
    currentItems = itemsCache || [];
    renderItemsGrid(currentItems);
  } catch (e) {
    document.getElementById('items-container').innerHTML = dogError('Couldn\'t load items. Check your connection.');
  }
}

function isNewItem(item) {
  if (!item.CreatedAt) return false;
  const created  = new Date(item.CreatedAt);
  const diffDays = (new Date() - created) / (1000 * 60 * 60 * 24);
  return diffDays <= 60;
}

function newBadgeHtml() {
  return '<div class="new-item-banner">✦ NEW</div>';
}

function renderItemsGrid(items) {
  const container = document.getElementById('items-container');
  if (!items || items.length === 0) {
    container.innerHTML = `
      <div class="dog-state">
        ${dogEmpty('No designs yet')}
        <button class="btn btn-primary" style="margin-top:1rem" onclick="window._cameFromInventory=true;document.querySelector('[data-page=new-item-design]').click()">✚ Add Design</button>
      </div>`;
    return;
  }
  container.innerHTML = `<div class="items-grid">${items.map(item => `
    <div class="item-card ${item.Status === 'Retired' ? 'item-retired' : ''}" onclick="openItemDetail('${item.ItemID}')">
      <div class="item-card-image">
        ${item.Photo
          ? `<img src="${fixPhotoUrl(item.Photo)}" alt="${item.DisplayName || item.Name}" loading="lazy">`
          : `<div class="item-card-placeholder"><svg width="36" height="36" viewBox="0 0 36 36" fill="none"><rect x="3" y="6" width="30" height="24" rx="3" stroke="#C4A882" stroke-width="2"/><circle cx="12" cy="14" r="3" stroke="#C4A882" stroke-width="2"/><path d="M3 25l8-7 6 5 4-4 12 9" stroke="#C4A882" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg></div>`}
        ${isNewItem(item) ? newBadgeHtml() : ''}
      </div>
      <div class="item-card-body">
        <div class="item-card-id">#${item.ItemID}</div>
        <div class="item-card-name">${item.DisplayName || item.Name || 'Untitled'}</div>
        <div class="item-card-meta">
          <span class="item-card-price">$${Number(item.UnitPrice || 0).toFixed(2)}</span>
          <span class="badge ${item.Status === 'Retired' ? 'badge-coral' : item.Status === 'Limited' ? 'badge-amber' : 'badge-green'}">
            ${item.Status || 'Open'}
          </span>
        </div>
      </div>
    </div>`).join('')}</div>`;
}

function renderItemsList(items) {
  const container = document.getElementById('items-container');
  if (!items || items.length === 0) { renderItemsGrid(items); return; }
  container.innerHTML = `
    <div class="card" style="padding:0;overflow:hidden;">
      <table class="items-table">
        <thead><tr><th>Image</th><th>ID</th><th>Name</th><th>Type</th><th>Price</th><th>Status</th></tr></thead>
        <tbody>${items.map(item => `
          <tr onclick="openItemDetail('${item.ItemID}')">
            <td><div class="table-thumbnail">${item.Photo ? `<img src="${fixPhotoUrl(item.Photo)}" alt="" style="width:100%;height:100%;object-fit:cover;">` : '<svg width="22" height="22" viewBox="0 0 36 36" fill="none"><rect x="3" y="6" width="30" height="24" rx="3" stroke="#C4A882" stroke-width="2"/></svg>'}</div></td>
            <td><span style="color:var(--teal);font-weight:700;font-size:0.8rem;">#${item.ItemID}</span></td>
            <td><strong>${item.DisplayName || item.Name || 'Untitled'}</strong></td>
            <td style="color:var(--brown-light)">${item.ProductType || '—'}</td>
            <td><strong>$${Number(item.UnitPrice || 0).toFixed(2)}</strong></td>
            <td><span class="badge ${item.Status === 'Retired' ? 'badge-coral' : item.Status === 'Limited' ? 'badge-amber' : 'badge-green'}">${item.Status || 'Open'}</span></td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

async function openItemDetail(itemId) {
  const item = (itemsCache || []).find(i => String(i.ItemID) === String(itemId));
  if (!item) return;
  if (!tagsCache) {
    try { const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getTags`); const d = await r.json(); if (d.success) tagsCache = d.tags; } catch (e) {}
  }
  openDetailPanel(`
    <div class="detail-id">#${item.ItemID}</div>
    <div class="detail-name">${item.DisplayName || item.Name || 'Untitled'}</div>
    <div class="detail-image">
      ${item.Photo ? `<img src="${fixPhotoUrl(item.Photo)}" alt="${item.DisplayName}">` : `<div class="detail-image-placeholder">${dogEmpty()}</div>`}
    </div>
    <div class="detail-stats">
      <div class="detail-stat"><div class="detail-stat-label">Retail Price</div><div class="detail-stat-value">$${Number(item.UnitPrice || 0).toFixed(2)}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Type</div><div class="detail-stat-value" style="font-size:1rem">${item.ProductType || '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">On Hand</div><div class="detail-stat-value">${item.StartingAtHome || 0}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">At Stores</div><div class="detail-stat-value" id="detail-at-stores-${item.ItemID}" style="color:var(--teal)">...</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Total Printed</div><div class="detail-stat-value" id="detail-printed-${item.ItemID}" style="color:var(--amber)">...</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Status</div><div class="detail-stat-value" style="font-size:1rem">${item.Status === 'Retired' ? '🪦 Retired' : item.Status === 'Limited' ? '⭐ Limited' : '✅ Open'}</div></div>
    </div>
    ${item.Notes ? `<div style="margin-bottom:1rem;font-size:0.875rem;color:var(--brown-mid);background:var(--cream);padding:0.75rem;border-radius:var(--radius-sm);">📝 ${item.Notes}</div>` : ''}
    ${(item._tagIds && item._tagIds.length > 0) ? `<div style="margin-bottom:1rem;"><div style="font-size:0.75rem;font-weight:600;color:var(--brown-light);text-transform:uppercase;letter-spacing:0.04em;margin-bottom:0.4rem;">Tags</div><div style="display:flex;flex-wrap:wrap;gap:0.35rem;">${item._tagIds.map(id => { const t = (tagsCache || []).find(t => t.TagID === id); return t ? `<span style="display:inline-block;padding:0.2rem 0.6rem;font-size:0.75rem;background:var(--cream);border:1px solid var(--brown-lightest);border-radius:999px;color:var(--brown-mid);">${t.TagName}</span>` : ''; }).join('')}</div></div>` : ''}
    <hr class="detail-divider">
    <div class="detail-actions">
      <button class="btn btn-primary" onclick="openEditItemModal('${item.ItemID}')">✏️ Edit Design</button>
      <button class="btn btn-secondary" onclick="closeDetailPanel();document.querySelector('[data-page=print-stock-updater]').click()">🖨️ Log Print Run</button>
    </div>
    <div style="position:sticky;bottom:0;background:var(--white);display:flex;justify-content:space-between;align-items:center;padding:0.75rem 0;margin-top:1.5rem;border-top:1px solid var(--border);">
      <button class="btn btn-secondary btn-sm" onclick="navigateItemDetail('${item.ItemID}', -1)">← Prev</button>
      <span style="font-size:0.75rem;color:var(--brown-light);">Browse designs</span>
      <button class="btn btn-secondary btn-sm" onclick="navigateItemDetail('${item.ItemID}', 1)">Next →</button>
    </div>`);
  // Load consignment and print run data for this item
  try {
    const [ctData, prData] = await Promise.all([
      window._consignmentTotals ? Promise.resolve(window._consignmentTotals) : fetch(`${GOOGLE_SCRIPT_URL}?action=getConsignmentTotals`).then(r => r.json()).then(d => { if (d.success) { window._consignmentTotals = d.totals; return d.totals; } return {}; }),
      window._printRunTotals ? Promise.resolve(window._printRunTotals) : fetch(`${GOOGLE_SCRIPT_URL}?action=getPrintRunTotals`).then(r => r.json()).then(d => { if (d.success) { window._printRunTotals = d.totals; return d.totals; } return {}; }),
    ]);
    const storesEl = document.getElementById(`detail-at-stores-${item.ItemID}`);
    if (storesEl) storesEl.textContent = ctData[String(item.ItemID)] || 0;
    const printedEl = document.getElementById(`detail-printed-${item.ItemID}`);
    if (printedEl) printedEl.textContent = prData[String(item.ItemID)] || 0;
  } catch(e) {
    const storesEl = document.getElementById(`detail-at-stores-${item.ItemID}`);
    if (storesEl) storesEl.textContent = '—';
    const printedEl = document.getElementById(`detail-printed-${item.ItemID}`);
    if (printedEl) printedEl.textContent = '—';
  }
}

function navigateItemDetail(currentId, direction) {
  const items = itemsCache || [];
  if (items.length === 0) return;
  const currentIdx = items.findIndex(i => String(i.ItemID) === String(currentId));
  if (currentIdx === -1) return;
  let nextIdx = currentIdx + direction;
  if (nextIdx < 0) nextIdx = items.length - 1;
  if (nextIdx >= items.length) nextIdx = 0;
  openItemDetail(items[nextIdx].ItemID);
}

function openEditItemModal(itemId) {
  const item = (itemsCache || []).find(i => String(i.ItemID) === String(itemId));
  if (!item) return;
  openModal(`
    <div class="modal-title">✏️ Edit Design #${item.ItemID}</div>
    <form id="edit-item-form">
      <div class="form-field">
        <label class="field-label">Design Name</label>
        <input class="field-input" type="text" name="designName" value="${item.Name || ''}" required>
      </div>
      <div class="form-grid">
        <div class="form-field">
          <label class="field-label">Card Type</label>
          <input class="field-input" type="text" name="itemType" value="${item.ProductType || ''}">
        </div>
        <div class="form-field">
          <label class="field-label">Retail Price</label>
          <input class="field-input" type="number" name="unitPrice" value="${item.UnitPrice || ''}" step="0.01">
        </div>
      </div>
      <div class="form-field">
        <label class="field-label">Photo</label>
        <div class="photo-upload-row">
          <input class="field-input" type="text" name="photo" id="edit-photo-url" value="${item.Photo || ''}" placeholder="Paste URL or upload below..." onchange="handlePhotoUrlPaste(this, '.photo-preview-wrap')" onpaste="setTimeout(() => handlePhotoUrlPaste(this, '.photo-preview-wrap'), 100)">
        </div>
        ${item.Photo ? `<div class="photo-preview-wrap"><img class="photo-preview-thumb" src="${fixPhotoUrl(item.Photo)}" alt="Current photo"></div>` : ''}
        <div class="photo-upload-area" onclick="document.getElementById('edit-photo-file').click()">
          <input type="file" id="edit-photo-file" accept="image/*" style="display:none" onchange="handlePhotoUpload(this)">
          <div class="photo-upload-icon">📷</div>
          <div class="photo-upload-text">Tap to take photo or choose from device</div>
        </div>
        <div id="photo-upload-status" style="font-size:0.75rem;color:var(--teal);margin-top:0.3rem;display:none;"></div>
      </div>
      <div class="form-field">
        <label class="field-label">Notes</label>
        <input class="field-input" type="text" name="notes" value="${item.Notes || ''}">
      </div>
      <div class="form-field">
        <label class="field-label">Tags</label>
        <div id="edit-tags-container">${dogLoading('Loading tags...')}</div>
      </div>
      <div class="form-field">
        <label class="field-label">Status</label>
        <select class="field-input" name="status">
          <option value="Open" ${(!item.Status || item.Status === 'Open') ? 'selected' : ''}>Open</option>
          <option value="Limited" ${item.Status === 'Limited' ? 'selected' : ''}>Limited</option>
          <option value="Retired" ${item.Status === 'Retired' ? 'selected' : ''}>Retired</option>
        </select>
      </div>
      <div id="edit-item-btn-wrap"><button type="submit" class="btn btn-primary" style="width:100%;margin-top:0.5rem">Save Changes</button></div>
      <div id="edit-item-status" class="form-status"></div>
    </form>`);

  loadTagsForEdit('edit-tags-container', item._tagIds || []);
  document.getElementById('edit-item-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status       = document.getElementById('edit-item-status');
    const btnWrap      = document.getElementById('edit-item-btn-wrap');
    const fd           = new FormData(e.target);
    const selectedTags = [...document.querySelectorAll('.edit-tag-check:checked')].map(c => c.value);
    const payload      = {
      action: 'updateItem',
      itemData: {
        itemId, designName: fd.get('designName'), itemType: fd.get('itemType'),
        unitPrice: fd.get('unitPrice'), photo: fixPhotoUrl(fd.get('photo')),
        notes: fd.get('notes'), status: fd.get('status'), tags: selectedTags,
      }
    };
    if (btnWrap) btnWrap.style.display = 'none';
    status.className = 'form-status loading';
    status.textContent = 'Saving...';
    try {
      const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
      const result = await r.json();
      if (result.success) {
        itemsCache = null;
        status.className  = 'form-status success';
        status.textContent = '✅ Design updated!';
        showToast('Design updated!', 'success');
        setTimeout(async () => { closeModal(); itemsCache = null; await fetchAndDisplayItems(); openItemDetail(itemId); }, 1200);
      } else throw new Error(result.error);
    } catch (err) {
      status.className  = 'form-status error';
      status.textContent = '❌ ' + err.message;
      if (btnWrap) btnWrap.style.display = '';
    }
  });
}

async function handlePhotoUpload(input) {
  const file     = input.files[0];
  if (!file) return;
  const statusEl  = document.getElementById('photo-upload-status');
  const urlInput  = document.getElementById('edit-photo-url');
  if (statusEl) { statusEl.style.display = 'block'; statusEl.style.color = 'var(--teal)'; statusEl.textContent = 'Processing photo...'; }
  const canvas = document.createElement('canvas');
  const img    = new Image();
  const reader = new FileReader();
  reader.onload = (e) => {
    img.onload = async () => {
      const maxSize = 800;
      let w = img.width, h = img.height;
      if (w > h && w > maxSize) { h = Math.round(h * maxSize / w); w = maxSize; }
      else if (h > maxSize) { w = Math.round(w * maxSize / h); h = maxSize; }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      const base64 = canvas.toDataURL('image/jpeg', 0.75).split(',')[1];
      if (statusEl) statusEl.textContent = 'Uploading to Drive...';
      try {
        const controller = new AbortController();
        const timeout    = setTimeout(() => controller.abort(), 30000);
        const r          = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'uploadPhoto', base64, filename: file.name }), signal: controller.signal });
        clearTimeout(timeout);
        const result = await r.json();
        if (result.success && result.url) {
          if (urlInput) urlInput.value = result.url;
          // Show a live preview on success
          let previewWrap = document.querySelector('.photo-preview-wrap');
          if (!previewWrap) {
            previewWrap = document.createElement('div');
            previewWrap.className = 'photo-preview-wrap';
            document.querySelector('.photo-upload-area')?.before(previewWrap);
          }
          previewWrap.innerHTML = `<img class="photo-preview-thumb" src="${fixPhotoUrl(result.url)}" alt="Uploaded photo">`;
          if (statusEl) { statusEl.textContent = '✅ Photo uploaded!'; statusEl.style.color = 'var(--green)'; }
        } else throw new Error(result.error || 'Upload failed');
      } catch (err) {
        const msg = err.name === 'AbortError' ? 'Upload timed out — check your connection.' : 'Drive upload failed: ' + err.message;
        if (statusEl) { statusEl.textContent = '⚠️ ' + msg; statusEl.style.color = 'var(--amber)'; }
      }
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

async function handleAddPhotoUpload(input) {
  const file     = input.files[0];
  if (!file) return;
  const statusEl = document.getElementById('add-photo-status');
  const urlInput = document.getElementById('add-photo-url');
  if (statusEl) { statusEl.style.display = 'block'; statusEl.textContent = 'Processing...'; statusEl.style.color = 'var(--teal)'; }
  const canvas = document.createElement('canvas');
  const img    = new Image();
  const reader = new FileReader();
  reader.onload = (e) => {
    img.onload = async () => {
      const maxSize = 800;
      let w = img.width, h = img.height;
      if (w > h && w > maxSize) { h = Math.round(h * maxSize / w); w = maxSize; }
      else if (h > maxSize) { w = Math.round(w * maxSize / h); h = maxSize; }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      const base64 = canvas.toDataURL('image/jpeg', 0.75).split(',')[1];
      try {
        const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'uploadPhoto', base64, filename: file.name }) });
        const result = await r.json();
        if (result.success && result.url) {
          if (urlInput) urlInput.value = result.url;
          // Show preview
          let previewWrap = document.querySelector('.add-photo-preview-wrap');
          if (!previewWrap) {
            previewWrap = document.createElement('div');
            previewWrap.className = 'add-photo-preview-wrap photo-preview-wrap';
            document.querySelector('#add-photo-url')?.before(previewWrap);
          }
          previewWrap.innerHTML = `<img class="photo-preview-thumb" src="${fixPhotoUrl(result.url)}" alt="Uploaded photo">`;
          if (statusEl) { statusEl.textContent = '✅ Photo ready!'; statusEl.style.color = 'var(--green)'; }
        } else throw new Error(result.error);
      } catch (err) {
        if (statusEl) { statusEl.textContent = '⚠️ Drive upload unavailable.'; statusEl.style.color = 'var(--amber)'; }
      }
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

async function handleMachinePhotoUpload(input) {
  const file = input.files[0];
  if (!file) return;
  const statusEl = document.getElementById('machine-photo-status');
  const urlInput = document.getElementById('machine-photo-url');
  if (statusEl) { statusEl.style.display = 'block'; statusEl.textContent = 'Processing...'; statusEl.style.color = 'var(--teal)'; }
  const canvas = document.createElement('canvas');
  const img = new Image();
  const reader = new FileReader();
  reader.onload = (e) => {
    img.onload = async () => {
      const maxSize = 800;
      let w = img.width, h = img.height;
      if (w > h && w > maxSize) { h = Math.round(h * maxSize / w); w = maxSize; }
      else if (h > maxSize) { w = Math.round(w * maxSize / h); h = maxSize; }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      const base64 = canvas.toDataURL('image/jpeg', 0.75).split(',')[1];
      if (statusEl) statusEl.textContent = 'Uploading to Drive...';
      try {
        const r = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'uploadPhoto', base64, filename: file.name }) });
        const result = await r.json();
        if (result.success && result.url) {
          if (urlInput) urlInput.value = result.url;
          let previewWrap = document.querySelector('.photo-preview-wrap');
          if (!previewWrap) {
            previewWrap = document.createElement('div');
            previewWrap.className = 'photo-preview-wrap';
            document.querySelector('.photo-upload-area')?.before(previewWrap);
          }
          previewWrap.innerHTML = `<img class="photo-preview-thumb" src="${fixPhotoUrl(result.url)}" alt="Uploaded photo">`;
          if (statusEl) { statusEl.textContent = '✅ Photo uploaded!'; statusEl.style.color = 'var(--green)'; }
        } else throw new Error(result.error || 'Upload failed');
      } catch (err) {
        if (statusEl) { statusEl.textContent = '⚠️ Upload failed: ' + err.message; statusEl.style.color = 'var(--amber)'; }
      }
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

async function loadTagsForEdit(containerId, selectedTagIds) {
  const container = document.getElementById(containerId);
  if (!container) return;
  if (!tagsCache) {
    try {
      const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getTags`);
      const d = await r.json();
      if (d.success) tagsCache = d.tags;
    } catch (e) { container.innerHTML = ''; return; }
  }
  const categories = {};
  (tagsCache || []).forEach(tag => {
    if (!categories[tag.Category]) categories[tag.Category] = [];
    categories[tag.Category].push(tag);
  });
  container.innerHTML = Object.entries(categories).map(([cat, tags]) => `
    <div class="tag-category-group">
      <div class="tag-category-label">${cat}</div>
      <div class="tag-chips-edit">
        ${tags.map(tag => `
          <label class="tag-chip-label">
            <input type="checkbox" class="edit-tag-check" value="${tag.TagID}" ${selectedTagIds.includes(tag.TagID) ? 'checked' : ''}>
            <span class="tag-chip-toggle">${tag.TagName}</span>
          </label>`).join('')}
      </div>
    </div>`).join('') + `
    <button type="button" class="btn btn-secondary btn-sm" style="margin-top:0.5rem" onclick="openAddTagModal()">+ Add New Tag</button>`;
}

function openAddTagModal() {
  openSubModal(`
    <div class="modal-title">+ Add New Tag</div>
    <form id="add-tag-form">
      <div class="form-field">
        <label class="field-label">Tag Name</label>
        <input class="field-input" type="text" name="tagName" placeholder="e.g. Beach & Ocean" required>
      </div>
      <div class="form-field">
        <label class="field-label">Category</label>
        <select class="field-input" name="category">
          <option>Theme</option><option>Audience</option><option>Rating</option><option>Format</option>
        </select>
      </div>
      <button type="submit" class="btn btn-primary" style="width:100%;margin-top:0.5rem">Add Tag</button>
      <div id="add-tag-status" class="form-status"></div>
    </form>`);

  document.getElementById('add-tag-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const fd     = new FormData(e.target);
    const status = document.getElementById('add-tag-status');
    status.className  = 'form-status loading';
    status.textContent = 'Adding...';
    try {
      const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'addTag', tagData: { tagName: fd.get('tagName'), category: fd.get('category') } }) });
      const result = await r.json();
      if (result.success) {
        tagsCache = null;
        status.className  = 'form-status success';
        status.textContent = '✅ Tag added!';
        // Reload tags in the edit form if it's open
        const editContainer = document.getElementById('edit-tags-container');
        if (editContainer) {
          const currentlyChecked = [...document.querySelectorAll('.edit-tag-check:checked')].map(c => c.value);
          loadTagsForEdit('edit-tags-container', currentlyChecked);
        }
        const newContainer = document.getElementById('new-tags-container');
        if (newContainer) {
          const currentlyChecked = [...document.querySelectorAll('.new-tag-check:checked')].map(c => c.value);
          loadTagsForEdit('new-tags-container', currentlyChecked);
        }
        setTimeout(() => closeSubModal(), 1000);
      } else throw new Error(result.error);
    } catch (err) {
      status.className  = 'form-status error';
      status.textContent = '❌ ' + err.message;
    }
  });
}

async function openManageTagsModal() {
  openModal(`<div class="modal-title">Manage Tags</div><div id="manage-tags-body">${dogLoading('Loading tags...')}</div>`);
  if (!tagsCache) {
    try { const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getTags`); const d = await r.json(); if (d.success) tagsCache = d.tags; } catch (e) {}
  }
  renderManageTagsList();
}

function renderManageTagsList() {
  const body = document.getElementById('manage-tags-body');
  if (!body || !tagsCache) return;
  const categories = {};
  (tagsCache || []).forEach(tag => {
    if (!categories[tag.Category]) categories[tag.Category] = [];
    categories[tag.Category].push(tag);
  });
  body.innerHTML = Object.entries(categories).map(([cat, tags]) => `
    <div style="margin-bottom:1rem;">
      <div style="font-size:0.75rem;font-weight:600;color:var(--brown-light);text-transform:uppercase;letter-spacing:0.04em;margin-bottom:0.4rem;">${cat}</div>
      <div style="display:flex;flex-wrap:wrap;gap:0.35rem;">
        ${tags.map(tag => `
          <span style="display:inline-flex;align-items:center;gap:0.3rem;padding:0.25rem 0.6rem;font-size:0.8rem;background:var(--cream);border:1px solid var(--brown-lightest);border-radius:999px;color:var(--brown-mid);">
            ${tag.TagName}
            <button onclick="deleteTagFromSystem('${tag.TagID}')" style="background:none;border:none;color:var(--coral);cursor:pointer;font-size:0.9rem;padding:0;line-height:1;" title="Delete tag">✕</button>
          </span>`).join('')}
      </div>
    </div>`).join('') + `
    <button class="btn btn-secondary btn-sm" style="margin-top:0.5rem;" onclick="openAddTagModal()">+ Add New Tag</button>`;
}

async function deleteTagFromSystem(tagId) {
  const tag = (tagsCache || []).find(t => t.TagID === tagId);
  if (!confirm('Delete tag "' + (tag?.TagName || tagId) + '"? This removes it from all designs.')) return;
  try {
    const r = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'deleteTag', tagId }) });
    const result = await r.json();
    if (result.success) {
      tagsCache = tagsCache.filter(t => t.TagID !== tagId);
      renderManageTagsList();
      showToast('Tag deleted', 'success');
    } else throw new Error(result.error);
  } catch (e) {
    showToast('Error: ' + e.message, '');
  }
}

// ============================================================
// ADD NEW DESIGN
// ============================================================
async function renderNewItemDesignPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Add New Design</h1><p class="page-subtitle">Register a new card or print design</p></div>
    </div>
    <div class="form-page-container">${dogLoading('Loading...')}</div>`;

  try {
    const [idRes, tagsRes, typesRes] = await Promise.all([
      fetch(`${GOOGLE_SCRIPT_URL}?action=getNextItemId`).then(r => r.json()),
      fetch(`${GOOGLE_SCRIPT_URL}?action=getTags`).then(r => r.json()),
      fetch(`${GOOGLE_SCRIPT_URL}?action=getProductTypes`).then(r => r.json()),
    ]);
    // Bail out if user navigated away while fetching
    if (sessionStorage.getItem('pba_current_page') !== 'new-item-design') return;
    if (!idRes.success) throw new Error(idRes.error);
    if (tagsRes.success) tagsCache = tagsRes.tags;
    const productTypes = typesRes.success ? typesRes.types : [];
    const nextId = idRes.nextId;
    const today  = new Date().toLocaleDateString('en-CA');
    const categories = {};
    (tagsCache || []).forEach(tag => {
      if (!categories[tag.Category]) categories[tag.Category] = [];
      categories[tag.Category].push(tag);
    });

    appContainer.innerHTML = `
      <div class="page-header">
        <div><h1 class="page-title">Add New Design</h1><p class="page-subtitle">Register a new card or print design</p></div>
      </div>
      <div class="form-page-container">
        <form id="new-item-form">
          <div class="form-section">
            <div class="form-section-title">Design Information</div>
            <div class="form-grid">
              <div class="form-field">
                <label class="field-label">Design Number</label>
                <input class="field-input" type="text" name="itemId" value="${nextId}" readonly>
                <div class="field-hint">Auto-generated — next in sequence</div>
              </div>
              <div class="form-field">
                <label class="field-label">Date Added</label>
                <input class="field-input" type="text" name="dateAdded" value="${today}" readonly>
              </div>
            </div>
            <div class="form-field">
              <label class="field-label">Design Name *</label>
              <input class="field-input" type="text" name="designName" placeholder="e.g. Gulf Coast Sunrise" required>
            </div>
            <div class="form-field">
              <label class="field-label">Card / Item Type *</label>
              <select class="field-input" name="itemType" id="new-item-type" required>
                <option value="" disabled selected>Select a type...</option>
                ${productTypes.map(t => `<option value="${t.TypeName || t.Name}" data-price="${t.DefaultRetailPrice || t.RetailPrice || t.Price || ''}">${t.TypeName || t.Name}</option>`).join('')}
                <option value="__new__">+ Add New Type...</option>
              </select>
              <div id="new-type-fields" style="display:none;margin-top:0.5rem;">
                <input class="field-input" type="text" id="new-type-name" placeholder="New type name..." style="margin-bottom:0.4rem;">
                <input class="field-input" type="number" id="new-type-price" placeholder="Retail price (e.g. 6.00)" step="0.01" min="0">
              </div>
            </div>
            <div class="form-field">
              <label class="field-label">Retail Price</label>
              <input class="field-input" type="number" name="unitPrice" id="new-item-price" placeholder="e.g. 6.00" step="0.01" min="0" readonly style="background:var(--cream);color:var(--brown-mid);">
            </div>
            <div class="form-field">
              <label class="field-label">Status</label>
              <select class="field-input" name="status">
                <option value="Open" selected>Open</option>
                <option value="Limited">Limited</option>
              </select>
            </div>
            <div class="form-field">
              <label class="field-label">First Print Run Quantity *</label>
              <input class="field-input" type="number" name="firstRun" placeholder="e.g. 50" min="0" required>
            </div>
          </div>
          <div class="form-section">
            <div class="form-section-title">Tags & Categories</div>
            ${Object.entries(categories).filter(([cat]) => cat !== 'Format').map(([cat, tags]) => `
              <div class="tag-category-group">
                <div class="tag-category-label">${cat}</div>
                <div class="tag-chips-edit">
                  ${tags.map(tag => `
                    <label class="tag-chip-label">
                      <input type="checkbox" class="new-tag-check" value="${tag.TagID}">
                      <span class="tag-chip-toggle">${tag.TagName}</span>
                    </label>`).join('')}
                </div>
              </div>`).join('')}
            ${(categories['Format'] || []).map(tag => `<input type="checkbox" class="new-tag-check" value="${tag.TagID}" data-format-tag="${tag.TagName}" style="display:none">`).join('')}
            <button type="button" class="btn btn-secondary btn-sm" style="margin-top:0.5rem" onclick="openAddTagModal()">+ Add New Tag</button>
          </div>
          <div class="form-section">
            <div class="form-section-title">Photo URL (Optional)</div>
            <div class="form-field">
              <div class="photo-upload-area" onclick="document.getElementById('add-photo-file').click()" style="margin-bottom:0.5rem;">
                <input type="file" id="add-photo-file" accept="image/*" style="display:none" onchange="handleAddPhotoUpload(this)">
                <div class="photo-upload-icon">📷</div>
                <div class="photo-upload-text">Tap to take photo or choose from device</div>
              </div>
              <input class="field-input" type="text" name="photo" id="add-photo-url" placeholder="Or paste a URL..." onchange="handlePhotoUrlPaste(this, '.add-photo-preview-wrap')" onpaste="setTimeout(() => handlePhotoUrlPaste(this, '.add-photo-preview-wrap'), 100)">
              <div id="add-photo-status" style="font-size:0.75rem;margin-top:0.3rem;display:none;"></div>
            </div>
          </div>
          <div class="form-field">
            <label class="field-label">Notes</label>
            <input class="field-input" type="text" name="notes" placeholder="Any additional notes">
          </div>
          <div id="add-design-btn-wrap"><button type="submit" class="btn btn-primary btn-lg" style="width:100%">✚ Add This Design</button></div>
          <div id="form-status" class="form-status"></div>
        </form>
      </div>`;

    document.getElementById('new-item-form').addEventListener('submit', handleAddNewItem);
    document.getElementById('new-item-type').addEventListener('change', (e) => {
      const selected = e.target.selectedOptions[0];
      const priceInput = document.getElementById('new-item-price');
      const newTypeFields = document.getElementById('new-type-fields');

      // Handle "+ Add New Type..."
      if (e.target.value === '__new__') {
        newTypeFields.style.display = '';
        priceInput.value = '';
        priceInput.readOnly = false;
        priceInput.style.background = '';
        priceInput.style.color = '';
        document.getElementById('new-type-name').focus();
        // When new type price is entered, sync to main price field
        document.getElementById('new-type-price').oninput = function() {
          priceInput.value = this.value ? parseFloat(this.value).toFixed(2) : '';
        };
        return;
      }
      newTypeFields.style.display = 'none';

      // Auto-fill price
      const price = selected?.dataset.price;
      if (price) {
        priceInput.value = parseFloat(price).toFixed(2);
        priceInput.readOnly = true;
        priceInput.style.background = 'var(--cream)';
        priceInput.style.color = 'var(--brown-mid)';
      } else {
        priceInput.value = '';
        priceInput.readOnly = false;
        priceInput.style.background = '';
        priceInput.style.color = '';
      }

      // Auto-check matching Format tag, uncheck others
      const typeName = e.target.value;
      document.querySelectorAll('[data-format-tag]').forEach(cb => {
        cb.checked = cb.dataset.formatTag === typeName;
      });
    });
  } catch (e) {
    appContainer.innerHTML = `<div class="dog-state">${dogError(e.message)}</div>`;
  }
}

async function handleAddNewItem(event) {
  event.preventDefault();
  const form     = event.target;
  const status   = document.getElementById('form-status');
  const btnWrap  = document.getElementById('add-design-btn-wrap');
  if (btnWrap) btnWrap.style.display = 'none';
  status.className  = 'form-status loading';
  status.textContent = 'Saving design...';
  const formData     = new FormData(form);
  const rawData      = Object.fromEntries(formData.entries());
  // Handle "Add New Type" — override itemType and price
  if (rawData.itemType === '__new__') {
    const newName = document.getElementById('new-type-name')?.value.trim();
    const newPrice = document.getElementById('new-type-price')?.value;
    if (!newName) { status.className = 'form-status error'; status.textContent = '❌ Please enter a name for the new type.'; if (btnWrap) btnWrap.style.display = ''; return; }
    rawData.itemType = newName;
    rawData.unitPrice = newPrice ? parseFloat(newPrice).toFixed(2) : rawData.unitPrice;
    // Save the new product type and create a matching Format tag
    fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'addProductType', typeData: { typeName: newName, retailPrice: rawData.unitPrice } }) });
    const tagRes = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'addTag', tagData: { tagName: newName, category: 'Format' } }) });
    const tagResult = await tagRes.json();
    if (tagResult.success && tagResult.tagId) {
      rawData._newFormatTagId = tagResult.tagId;
    }
    tagsCache = null;
  }
  // Auto-format price to 2 decimals
  if (rawData.unitPrice) rawData.unitPrice = parseFloat(rawData.unitPrice).toFixed(2);
  const selectedTags = [...document.querySelectorAll('.new-tag-check:checked')].map(c => c.value);
  // Include the new format tag if we just created one
  if (rawData._newFormatTagId && !selectedTags.includes(rawData._newFormatTagId)) {
    selectedTags.push(rawData._newFormatTagId);
  }
  const payload      = {
    action: 'addItem',
    itemData: { itemId: rawData.itemId, designName: rawData.designName, itemType: rawData.itemType, unitPrice: rawData.unitPrice, photo: fixPhotoUrl(rawData.photo), notes: rawData.notes, dateAdded: rawData.dateAdded, status: rawData.status || 'Open', tags: selectedTags },
    printRunData: { itemId: rawData.itemId, quantity: rawData.firstRun, date: rawData.dateAdded }
  };
  try {
    const response = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
    const result   = await response.json();
    if (result.success) {
      itemsCache = null;
      status.className  = 'form-status success';
      status.textContent = `✅ "${rawData.designName}" added successfully!`;
      showToast('Design added!', 'success');
      // If came from inventory page, go back there; otherwise reload add new design
      setTimeout(() => {
        if (sessionStorage.getItem('pba_current_page') !== 'new-item-design') return;
        if (window._cameFromInventory) {
          window._cameFromInventory = false;
          document.querySelector('[data-page="main-inventory"]')?.click();
        } else {
          window.scrollTo(0, 0);
          loadPage('new-item-design');
        }
      }, 1500);
    } else throw new Error(result.error);
  } catch (e) {
    status.className  = 'form-status error';
    status.textContent = `❌ Error: ${e.message}`;
    if (btnWrap) btnWrap.style.display = '';
  }
}

// ============================================================
// PRINT STOCK UPDATER
// ============================================================
async function renderPrintStockUpdaterPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Print Stock Update</h1><p class="page-subtitle">Log a new print run and update stock</p></div>
    </div>
    <div class="form-page-container">${dogLoading('Loading designs...')}</div>`;

  try {
    if (!itemsCache) {
      const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getItems`);
      const d = await r.json();
      if (!d.success) throw new Error(d.error);
      itemsCache = d.items.filter(i => i.ItemID);
    }
    const today = new Date().toLocaleDateString('en-CA');
    appContainer.innerHTML = `
      <div class="page-header">
        <div><h1 class="page-title">Print Stock Update</h1><p class="page-subtitle">Log a new print run and update stock</p></div>
      </div>
      <div class="form-page-container">
        <form id="print-stock-form">
          <div class="form-section">
            <div class="form-section-title">Print Run Details</div>
            <div class="form-field">
              <label class="field-label">Select Design *</label>
              <select id="item-select-stock" name="itemId" required>
                <option value="" disabled selected>Search or select a design...</option>
                ${itemsCache.map(i => `<option value="${i.ItemID}">${i.ItemID} — ${i.DisplayName || i.Name}</option>`).join('')}
              </select>
            </div>
            <div class="form-grid">
              <div class="form-field">
                <label class="field-label">Quantity Printed *</label>
                <input class="field-input" type="number" name="quantity" placeholder="e.g. 50" min="1" required>
              </div>
              <div class="form-field">
                <label class="field-label">Print Run Date *</label>
                <input class="field-input" type="date" name="printRunDate" value="${today}" required>
              </div>
            </div>
          </div>
          <div id="print-run-btn-wrap"><button type="submit" class="btn btn-primary btn-lg" style="width:100%">🖨️ Log Print Run</button></div>
          <div id="form-status" class="form-status"></div>
        </form>
      </div>`;

    new Choices('#item-select-stock', { searchEnabled: true, itemSelectText: '' });
    document.getElementById('print-stock-form').addEventListener('submit', handleAddPrintRun);
  } catch (e) {
    appContainer.innerHTML = `<div class="dog-state">${dogError(e.message)}</div>`;
  }
}

async function handleAddPrintRun(event) {
  event.preventDefault();
  const form     = event.target;
  const status   = document.getElementById('form-status');
  const btnWrap  = document.getElementById('print-run-btn-wrap');
  if (btnWrap) btnWrap.style.display = 'none';
  status.className  = 'form-status loading';
  status.textContent = 'Saving print run...';
  const formData = new FormData(form);
  const payload  = { action: 'addPrintRun', itemId: formData.get('itemId'), quantity: formData.get('quantity'), date: formData.get('printRunDate') };
  try {
    const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
    const result = await r.json();
    if (result.success) {
      itemsCache = null;
      window._printRunTotals = null;
      window._consignmentTotals = null;
      status.className  = 'form-status success';
      status.textContent = `✅ Print run logged for design #${payload.itemId}`;
      showToast('Print run saved!', 'success');
    } else throw new Error(result.error);
  } catch (e) {
    status.className  = 'form-status error';
    status.textContent = `❌ Error: ${e.message}`;
    if (btnWrap) btnWrap.style.display = '';
  }
}

// ============================================================
// RETAIL STOCK & SALES
// ============================================================
let retailInventoryState     = [];
let retailCurrentPartnerId   = null;
let retailCurrentPartnerName = null;

async function renderRecordSalePage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Retail Stock & Sales</h1><p class="page-subtitle">Update inventory at your retail partners</p></div>
    </div>
    <div id="rss-loading-state">${dogLoading('Loading partners...')}</div>
    <div id="rss-main" style="display:none;">
      <div class="card rss-partner-select-card" style="margin-bottom:1.5rem;">
        <div class="form-field">
          <label class="field-label">Select a Retail Partner</label>
          <select id="retail-partner-select"></select>
        </div>
      </div>
      <div id="retail-partner-info" style="display:none;"></div>
    </div>`;

  try {
    const [partnerRes, itemsRes] = await Promise.all([
      fetch(`${GOOGLE_SCRIPT_URL}?action=getRetailPartners`).then(r => r.json()),
      itemsCache ? Promise.resolve({ success: true, items: itemsCache }) :
        fetch(`${GOOGLE_SCRIPT_URL}?action=getItems`).then(r => r.json()),
    ]);
    if (!partnerRes.success) throw new Error(partnerRes.error);
    if (itemsRes.success) itemsCache = itemsRes.items.filter(i => i.ItemID);

    document.getElementById('rss-loading-state').style.display = 'none';
    document.getElementById('rss-main').style.display = 'block';

    const partnerSelect = document.getElementById('retail-partner-select');
    new Choices(partnerSelect, {
      choices: partnerRes.partners,
      searchEnabled: true,
      searchPlaceholderValue: 'Type to search partners...',
      itemSelectText: '',
      allowHTML: true,
      placeholderValue: 'Select a partner...',
      shouldSort: false,
    });

    partnerSelect.addEventListener('change', (e) => {
      const partnerId = e.target.value;
      if (!partnerId) { document.getElementById('retail-partner-info').style.display = 'none'; return; }
      const partner = (partnerRes.partners || []).find(p => p.value === partnerId);
      retailCurrentPartnerId   = partnerId;
      retailCurrentPartnerName = partner?.label || partnerId;
      loadPartnerInventoryView(partnerId, partner);
    });

  } catch (e) {
    document.getElementById('rss-loading-state').innerHTML = `<div class="dog-state">${dogError(e.message)}</div>`;
  }
}

async function loadPartnerInventoryView(partnerId, partnerMeta) {
  const infoArea = document.getElementById('retail-partner-info');
  infoArea.style.display = 'block';
  infoArea.innerHTML = dogLoading('Loading inventory...');

  try {
    const partnerRes = await fetch(`${GOOGLE_SCRIPT_URL}?action=getPartnerInventory&partnerId=${partnerId}`).then(r => r.json());
    if (!partnerRes.success) throw new Error(partnerRes.error);

    const { name, lastVisit, inventory } = partnerRes.data;
    retailCurrentPartnerName = name;

    // De-duplicate by designId (keep last entry) and filter out zero-stock
    const deduped = new Map();
    (inventory || []).forEach(item => deduped.set(String(item.designId), item));
    retailInventoryState = [...deduped.values()]
      .filter(item => (item.currentStock || 0) > 0)
      .map(item => ({
        designId:      item.designId,
        designName:    item.designName,
        unitPrice:     item.unitPrice || (itemsCache || []).find(i => String(i.ItemID) === String(item.designId))?.UnitPrice || 0,
        previousStock: item.currentStock,
        currentStock:  item.currentStock,
        pulled:        0,
        added:         0,
        isNew:         false,
      }));

    const skuCodesHtml = (() => { try { const codes = JSON.parse((partnerMeta && partnerMeta.retailItemCodes) || '[]'); return codes.length ? `<div style="display:flex;flex-wrap:wrap;gap:0.4rem;margin-top:0.5rem;">${codes.map(c => `<span style="background:var(--cream);border:1px solid var(--tan);border-radius:6px;padding:0.2rem 0.6rem;font-size:0.8rem;"><span style="color:var(--brown-mid)">${c.product}:</span> <span style="font-family:monospace;color:var(--teal);font-weight:600">${c.code}</span></span>`).join('')}</div>` : ''; } catch(e) { return ''; } })();

    infoArea.innerHTML = `
      <div class="partner-info-header rss-partner-header" style="margin-bottom:1rem;">
        <div>
          <h2 id="retail-partner-name" class="partner-name" style="font-size:1.75rem;">${name}</h2>
          <p style="font-size:0.95rem;margin-top:0.35rem;">Last Inventory: <span id="last-visit-date">${lastVisit || 'Never'}</span></p>
          ${skuCodesHtml}
        </div>
      </div>
      <div id="stock-summary-strip" class="card" style="padding:0.75rem 1rem;margin-bottom:1.25rem;"></div>
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0.75rem;flex-wrap:wrap;gap:0.5rem;">
        <h3 class="section-title" style="margin:0;">Update Inventory</h3>
        <div style="display:flex;gap:0.5rem;align-items:center;">
          <div class="view-toggle" title="Card size">
            <button class="view-toggle-btn ${retailCardSize === 'compact' ? 'active' : ''}" onclick="setRetailCardSize('compact',this)">▬</button>
            <button class="view-toggle-btn ${retailCardSize === 'comfy' ? 'active' : ''}" onclick="setRetailCardSize('comfy',this)">≡</button>
          </div>
          <button class="btn btn-secondary btn-sm" onclick="openAddDesignToPartnerModal()">✚ Add Design to Store</button>
          <button class="btn btn-primary" onclick="savePartnerInventory()">💾 Save All Updates</button>
        </div>
      </div>
      <div id="inventory-list-container"></div>
      <div style="margin-top:1rem;display:flex;justify-content:flex-end;">
        <button class="btn btn-primary" onclick="savePartnerInventory()">💾 Save All Updates</button>
      </div>
      <hr style="border:none;border-top:1px solid var(--tan);margin:2rem 0 1.5rem;">
      <h3 class="section-title" style="margin-bottom:0.75rem;">💰 Log Actual Sales</h3>
      <div class="card" style="padding:1rem;">
        <div class="form-grid">
          <div class="form-field">
            <label class="field-label">Month</label>
            <div style="display:flex;gap:0.4rem;">
              <select class="field-input" id="actual-sale-month-m" style="flex:1;">
                ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].map((m,i) => `<option value="${String(i+1).padStart(2,'0')}" ${i === new Date().getMonth() ? 'selected' : ''}>${m}</option>`).join('')}
              </select>
              <select class="field-input" id="actual-sale-month-y" style="width:5rem;">
                ${(() => { const y = new Date().getFullYear(); return [y-1,y].map(yr => `<option value="${yr}" ${yr===y?'selected':''}>${yr}</option>`).join(''); })()}
              </select>
            </div>
          </div>
          <div class="form-field">
            <label class="field-label">Actual Sales ($)</label>
            <input class="field-input" type="number" id="actual-sale-amount" step="0.01" min="0" placeholder="e.g. 125.50">
          </div>
          <div class="form-field">
            <label class="field-label">Cards Sold</label>
            <input class="field-input" type="number" id="actual-sale-cards" min="0" placeholder="e.g. 12">
          </div>
        </div>
        <div style="display:flex;gap:0.5rem;align-items:center;margin-top:0.5rem;">
          <button class="btn btn-primary btn-sm" onclick="submitActualSale()">Save Sale</button>
          <span id="actual-sale-status" style="font-size:0.8rem;"></span>
        </div>
        <div id="sales-history-container" style="margin-top:1rem;"></div>
      </div>`;

    renderInventoryCards();
    loadSalesHistory(partnerId);

  } catch (e) {
    infoArea.innerHTML = `<div class="dog-state">${dogError(e.message)}</div>`;
  }
}

async function submitActualSale() {
  const month = (document.getElementById('actual-sale-month-y')?.value || '') + '-' + (document.getElementById('actual-sale-month-m')?.value || '');
  const amount = document.getElementById('actual-sale-amount')?.value;
  const cardsSold = document.getElementById('actual-sale-cards')?.value;
  const status = document.getElementById('actual-sale-status');
  if (!month || !amount) { status.textContent = 'Please enter month and amount.'; status.style.color = 'var(--coral)'; return; }
  status.textContent = 'Saving...'; status.style.color = 'var(--teal)';
  try {
    const r = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({
      action: 'logActualSale',
      partnerId: retailCurrentPartnerId,
      partnerName: retailCurrentPartnerName,
      month: month,
      actualSales: parseFloat(amount),
      cardsSold: parseInt(cardsSold) || 0
    })});
    const result = await r.json();
    if (result.success) {
      status.textContent = '✅ Saved!'; status.style.color = 'var(--green)';
      document.getElementById('actual-sale-amount').value = '';
      document.getElementById('actual-sale-cards').value = '';
      loadSalesHistory(retailCurrentPartnerId);
    } else throw new Error(result.error);
  } catch (e) { status.textContent = '❌ ' + e.message; status.style.color = 'var(--coral)'; }
}

async function loadSalesHistory(partnerId) {
  const container = document.getElementById('sales-history-container');
  if (!container) return;
  try {
    const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getPartnerSalesHistory&partnerId=${partnerId}`);
    const data = await r.json();
    if (!data.success || !data.history || data.history.length === 0) {
      container.innerHTML = '<p style="font-size:0.8rem;color:var(--brown-mid);margin:0;">No sales logged yet.</p>';
      return;
    }
    container.innerHTML = `
      <div style="font-size:0.75rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--brown-mid);font-weight:700;margin-bottom:0.4rem;">Sales History</div>
      <div style="display:grid;grid-template-columns:1fr auto auto;gap:0.2rem 1rem;font-size:0.85rem;">
        ${data.history.map(h => `
          <span style="color:var(--brown-dark)">${h.month}</span>
          <span style="font-weight:600;color:var(--green);text-align:right;">$${Number(h.actualSales || 0).toFixed(2)}</span>
          <span style="color:var(--brown-mid);text-align:right;">${h.cardsSold ? h.cardsSold + ' cards' : '—'}</span>
        `).join('')}
      </div>`;
  } catch (e) {
    container.innerHTML = '';
  }
}

let retailSummaryExpanded = false;
let retailCardSize = window.innerWidth <= 768 ? 'compact' : 'comfy';

function toggleSummaryView() {
  retailSummaryExpanded = !retailSummaryExpanded;
  renderStockSummary();
}

function setRetailCardSize(mode, btn) {
  retailCardSize = mode;
  document.querySelectorAll('.view-toggle-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderInventoryCards();
}

function renderStockSummary() {
  const strip = document.getElementById('stock-summary-strip');
  if (!strip) return;

  if (retailInventoryState.length === 0) {
    strip.style.display = 'none';
    return;
  }
  strip.style.display = '';

  const totalCards = retailInventoryState.reduce((sum, item) => {
    return sum + Math.max(0, (item.currentStock || 0) + (item.added || 0) - (item.pulled || 0));
  }, 0);

  if (!retailSummaryExpanded) {
    // Compact: inline pills
    strip.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0.5rem;">
        <span style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--brown-mid);font-weight:700;">${retailInventoryState.length} Designs on Shelf</span>
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="font-size:0.8rem;font-weight:600;color:var(--teal);">${totalCards} total cards</span>
          <button class="btn btn-secondary btn-sm" style="padding:0.15rem 0.5rem;font-size:0.7rem;" onclick="toggleSummaryView()">Expand</button>
        </div>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:0.35rem;">
        ${retailInventoryState.map(item => {
          const stock = Math.max(0, (item.currentStock || 0) + (item.added || 0) - (item.pulled || 0));
          const shortName = (item.designName || '').replace(/^\d+\s*—\s*/, '');
          const displayName = shortName.length > 22 ? shortName.substring(0, 20) + '...' : shortName;
          return `<span style="display:inline-flex;align-items:center;gap:0.3rem;background:var(--cream);border:1px solid var(--border);border-radius:100px;padding:0.2rem 0.6rem 0.2rem 0.5rem;font-size:0.75rem;color:var(--brown-mid);white-space:nowrap;">
            <span style="background:var(--teal);color:white;border-radius:50%;width:1.2rem;height:1.2rem;display:inline-flex;align-items:center;justify-content:center;font-size:0.65rem;font-weight:700;flex-shrink:0;">${stock}</span>
            ${displayName}
          </span>`;
        }).join('')}
      </div>`;
  } else {
    // Expanded: larger cards in a grid
    strip.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0.75rem;">
        <span style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--brown-mid);font-weight:700;">${retailInventoryState.length} Designs on Shelf</span>
        <div style="display:flex;align-items:center;gap:0.5rem;">
          <span style="font-size:0.8rem;font-weight:600;color:var(--teal);">${totalCards} total cards</span>
          <button class="btn btn-secondary btn-sm" style="padding:0.15rem 0.5rem;font-size:0.7rem;" onclick="toggleSummaryView()">Compact</button>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(3, 1fr);gap:0.5rem;">
        ${retailInventoryState.map(item => {
          const stock = Math.max(0, (item.currentStock || 0) + (item.added || 0) - (item.pulled || 0));
          const shortName = (item.designName || '').replace(/^\d+\s*—\s*/, '');
          return `<div style="background:var(--cream);border:1px solid var(--border);border-radius:var(--radius-sm);padding:0.6rem 0.75rem;display:flex;align-items:center;gap:0.5rem;">
            <span style="background:var(--teal);color:white;border-radius:50%;width:1.6rem;height:1.6rem;display:inline-flex;align-items:center;justify-content:center;font-size:0.8rem;font-weight:700;flex-shrink:0;">${stock}</span>
            <div style="min-width:0;">
              <div style="font-size:0.85rem;font-weight:500;color:var(--brown);line-height:1.2;">${shortName}</div>
              <div style="font-size:0.7rem;color:var(--brown-light);">#${item.designId} · $${Number(item.unitPrice || 0).toFixed(2)}</div>
            </div>
          </div>`;
        }).join('')}
      </div>`;
  }
}

function renderInventoryCards() {
  const container = document.getElementById('inventory-list-container');
  if (!container) return;

  renderStockSummary();

  if (retailInventoryState.length === 0) {
    container.innerHTML = `<div class="dog-state">${dogEmpty('No inventory recorded for this partner yet.')}<br><button class="btn btn-primary" style="margin-top:1rem" onclick="openAddDesignToPartnerModal()">✚ Add First Design</button></div>`;
    return;
  }

  container.innerHTML = retailInventoryState.map((item, realIdx) => {
    const finalStock = Math.max(0, (item.currentStock || 0) + (item.added || 0) - (item.pulled || 0));
    const pendingRemoval = !item.isNew && finalStock === 0;
    const shortName = (item.designName || '').replace(/^\d+\s*—\s*/, '');
    const estSold = item.isNew ? 0 : Math.max(0, item.previousStock - item.currentStock);
    const estRevenue = (estSold * Number(item.unitPrice || 0)).toFixed(2);
    return `
      <div class="inventory-card ${retailCardSize === 'comfy' ? 'inventory-card-comfy' : ''} ${item.isNew ? 'inventory-card-new' : ''} ${pendingRemoval ? 'inventory-card-pending-removal' : ''}">
        <div class="inventory-card-info">
          <div class="design-name">${shortName}</div>
          <div class="design-id">#${item.designId} · $${Number(item.unitPrice || 0).toFixed(2)}</div>
          ${pendingRemoval ? `<div class="pending-removal-banner">⚠️ Will be removed on save</div>` : ''}
        </div>
        <div class="inventory-stock-col">
          <div class="inventory-stock-val" style="color:var(--amber);">${item.previousStock}</div>
          <div class="inventory-stock-label">Last Visit</div>
        </div>
        <div class="inventory-stock-col">
          <div class="inventory-stock-val new-stock-display" style="color:var(--teal);font-weight:700;">${(item.added > 0 || item.pulled > 0 || item.currentStock !== item.previousStock) ? finalStock : '—'}</div>
          <div class="inventory-stock-label">New Total</div>
        </div>
        <div class="inventory-actions">
          <div class="action-field">
            <label>New Count</label>
            <input type="number" class="action-input" min="0"
              oninput="updateInventoryField(${realIdx}, 'currentStock', this.value === '' ? '${item.previousStock}' : this.value)">
          </div>
          <div class="action-field">
            <label>Add +</label>
            <input type="number" class="action-input" placeholder="+" min="0"
              value="${item.added || ''}"
              oninput="updateInventoryField(${realIdx}, 'added', this.value)">
          </div>
          <div class="action-field">
            <label>Pull −</label>
            <input type="number" class="action-input" placeholder="−" min="0"
              value="${item.pulled || ''}"
              oninput="updateInventoryField(${realIdx}, 'pulled', this.value)">
          </div>
        </div>
        <div class="est-sold-row" style="font-size:0.7rem;color:var(--brown-light);text-align:right;min-width:65px;flex-shrink:0;">
          ${item.isNew ? '<em style="color:var(--teal);">New</em>' : ''}
        </div>
      </div>`;
  }).join('');
}

function updateInventoryField(idx, field, value) {
  if (!retailInventoryState[idx]) return;
  retailInventoryState[idx][field] = parseInt(value) || 0;

  const cards = document.querySelectorAll('.inventory-card');
  if (!cards[idx]) return;
  const item = retailInventoryState[idx];
  const finalStock = Math.max(0, (item.currentStock || 0) + (item.added || 0) - (item.pulled || 0));
  const pendingRemoval = !item.isNew && finalStock === 0;

  // Update "New Total" display — only show number if something changed
  const newStockEl = cards[idx].querySelector('.new-stock-display');
  const hasChanges = item.added > 0 || item.pulled > 0 || item.currentStock !== item.previousStock;
  if (newStockEl) newStockEl.textContent = hasChanges ? finalStock : '—';

  // Update estimated sold visibility
  const estSoldRow = cards[idx].querySelector('.est-sold-row');
  if (estSoldRow && !item.isNew) {
    const estSold = Math.max(0, item.previousStock - item.currentStock);
    const estRevenue = (estSold * Number(item.unitPrice || 0)).toFixed(2);
    estSoldRow.innerHTML = item.currentStock !== item.previousStock
      ? `Est. sold: <strong style="color:${estSold > 0 ? 'var(--green)' : 'var(--brown-light)'}">${estSold > 0 ? estSold : '—'}</strong>${estSold > 0 ? '<br><span style="color:var(--green);">$' + estRevenue + '</span>' : ''}`
      : '';
  }

  // Update estimated sold
  if (!item.isNew) {
    const estSold = Math.max(0, item.previousStock - item.currentStock);
    const estRevenue = (estSold * Number(item.unitPrice || 0)).toFixed(2);
    const estRow = cards[idx].querySelector('.est-sold-row');
    if (estRow) {
      estRow.innerHTML = `Est. sold: <strong style="color:${estSold > 0 ? 'var(--green)' : 'var(--brown-light)'}">${estSold > 0 ? estSold : '—'}</strong>${estSold > 0 ? '<br><span style="color:var(--green);">$' + estRevenue + '</span>' : ''}`;
    }
  }

  // Update pending-removal state
  if (pendingRemoval) {
    cards[idx].classList.add('inventory-card-pending-removal');
    if (!cards[idx].querySelector('.pending-removal-banner')) {
      const banner = document.createElement('div');
      banner.className = 'pending-removal-banner';
      banner.textContent = '⚠️ Will be removed on save';
      const info = cards[idx].querySelector('.inventory-card-info');
      if (info) info.appendChild(banner);
    }
  } else {
    cards[idx].classList.remove('inventory-card-pending-removal');
    cards[idx].querySelector('.pending-removal-banner')?.remove();
  }

  // Update summary strip
  renderStockSummary();
}

async function savePartnerInventory() {
  const btns = document.querySelectorAll('button[onclick="savePartnerInventory()"]');
  btns.forEach(b => { b.disabled = true; b.textContent = 'Saving...'; });

  try {
    const updates = retailInventoryState.map(item => {
      let finalStock = item.currentStock;
      if (item.added > 0)  finalStock += item.added;
      if (item.pulled > 0) finalStock = Math.max(0, finalStock - item.pulled);
      return {
        designId:      item.designId,
        designName:    item.designName,
        previousStock: item.previousStock,
        newStock:      finalStock,
        added:         item.added,
        pulled:        item.pulled,
        estimatedSold: Math.max(0, item.previousStock - finalStock),
        unitPrice:     item.unitPrice,
        isNew:         item.isNew,
      };
    });

    const today   = new Date().toLocaleDateString('en-CA');
    const payload = {
      action:      'updatePartnerInventory',
      partnerId:   retailCurrentPartnerId,
      partnerName: retailCurrentPartnerName,
      visitDate:   today,
      updates,
    };

    const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
    const result = await r.json();

    if (result.success) {
      // Clear caches so other pages reflect updated stock
      itemsCache = null;
      window._consignmentTotals = null;
      window._printRunTotals = null;
      showToast('✅ Inventory saved!', 'success');
      setTimeout(() => loadPartnerInventoryView(retailCurrentPartnerId, null), 800);
    } else throw new Error(result.error);
  } catch (e) {
    showToast('❌ Save failed: ' + e.message, '');
    btns.forEach(b => { b.disabled = false; b.textContent = '💾 Save All Updates'; });
  }
}

function openAddDesignToPartnerModal() {
  if (!itemsCache || itemsCache.length === 0) { showToast('Loading designs...', ''); return; }
  const existing  = new Set(retailInventoryState.map(i => String(i.designId)));
  const available = itemsCache.filter(i => !existing.has(String(i.ItemID)) && i.Status !== 'Retired');

  openModal(`
    <div class="modal-title">✚ Add Design to ${retailCurrentPartnerName}</div>
    <div class="form-field">
      <label class="field-label">Select Design *</label>
      <select class="field-input" id="add-partner-design-select">
        <option value="" disabled selected>Search or select...</option>
        ${available.map(i => `<option value="${i.ItemID}" data-price="${i.UnitPrice || 0}">${i.DisplayName || i.Name}</option>`).join('')}
      </select>
    </div>
    <div class="form-grid">
      <div class="form-field">
        <label class="field-label">Quantity Brought *</label>
        <input class="field-input" type="number" id="add-partner-qty" placeholder="e.g. 6" min="1">
      </div>
      <div class="form-field">
        <label class="field-label">Retail Price</label>
        <input class="field-input" type="number" id="add-partner-price" placeholder="Auto-filled" step="0.01" min="0" readonly style="background:var(--cream);color:var(--brown-mid);">
      </div>
    </div>
    <div id="add-partner-design-btn-wrap"><button class="btn btn-primary" style="width:100%;margin-top:0.5rem" onclick="confirmAddDesignToPartner()">✚ Add to Inventory</button></div>
    <div id="add-partner-design-status" class="form-status"></div>`);

  document.getElementById('add-partner-design-select')?.addEventListener('change', (e) => {
    const opt = e.target.selectedOptions[0];
    if (opt) document.getElementById('add-partner-price').value = opt.dataset.price || '';
  });
  new Choices('#add-partner-design-select', { searchEnabled: true, itemSelectText: '', searchPlaceholderValue: 'Type to search designs...' });
}

async function confirmAddDesignToPartner() {
  const selectEl = document.querySelector('#add-partner-design-select');
  const designId = selectEl?.value;
  const qty      = parseInt(document.getElementById('add-partner-qty')?.value);
  const price    = parseFloat(document.getElementById('add-partner-price')?.value) || 0;
  const status   = document.getElementById('add-partner-design-status');
  const btnWrap  = document.getElementById('add-partner-design-btn-wrap');

  if (!designId || isNaN(qty) || qty < 1) {
    status.className = 'form-status error'; status.textContent = '❌ Please select a design and enter a quantity.';
    return;
  }

  const item = (itemsCache || []).find(i => String(i.ItemID) === String(designId));
  const designName = item?.DisplayName || item?.Name || `Design #${designId}`;
  const unitPrice = price || item?.UnitPrice || 0;

  // Show saving state in modal
  if (btnWrap) btnWrap.style.display = 'none';
  status.className = 'form-status loading'; status.textContent = 'Adding to store...';

  // Auto-save just this new design to the system
  try {
    const today = new Date().toLocaleDateString('en-CA');
    const payload = {
      action: 'updatePartnerInventory',
      partnerId: retailCurrentPartnerId,
      partnerName: retailCurrentPartnerName,
      visitDate: today,
      updates: [{
        designId,
        designName,
        previousStock: 0,
        newStock: qty,
        added: qty,
        pulled: 0,
        estimatedSold: 0,
        unitPrice,
        isNew: true,
      }],
    };

    const r = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
    const result = await r.json();
    if (!result.success) throw new Error(result.error);

    // Add to local state so it shows on the page
    retailInventoryState.push({
      designId,
      designName,
      unitPrice,
      previousStock: 0,
      currentStock: qty,
      pulled: 0,
      added: 0, // reset since it's already saved
      isNew: false, // already persisted
    });

    closeModal();
    renderInventoryCards();
    showToast(`${designName} added and saved!`, 'success');
  } catch (e) {
    status.className = 'form-status error'; status.textContent = '❌ ' + e.message;
    if (btnWrap) btnWrap.style.display = '';
  }
}

// ============================================================
// RETAIL PARTNERS
// ============================================================
async function renderRetailPartnersPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Retail Partners</h1><p class="page-subtitle">Manage your consignment store relationships</p></div>
      <div class="page-actions"><button class="btn btn-primary" onclick="openAddPartnerModal()">✚ Add Partner</button></div>
    </div>
    <div class="items-toolbar" style="margin-bottom:1.25rem;">
      <div class="search-box"><span class="search-icon">⌕</span><input type="text" id="partners-search" placeholder="Search partners..."></div>
      <div class="view-toggle">
        <button class="view-toggle-btn active" id="partners-grid-btn" title="Tile view">⊞</button>
        <button class="view-toggle-btn" id="partners-list-btn" title="List view">☰</button>
      </div>
    </div>
    <div id="partners-container">${dogLoading('Loading partners...')}</div>`;

  try {
    const r    = await fetch(`${GOOGLE_SCRIPT_URL}?action=getRetailPartners`);
    const data = await r.json();
    if (!data.success) throw new Error(data.error);
    partnersCache = data.partners;

    document.getElementById('partners-search')?.addEventListener('input', (e) => {
      const q        = e.target.value.toLowerCase();
      const filtered = (partnersCache || []).filter(p => (p.label || '').toLowerCase().includes(q) || (p.city || '').toLowerCase().includes(q));
      const isGrid   = document.getElementById('partners-grid-btn')?.classList.contains('active');
      isGrid ? renderPartnersGrid(filtered) : renderPartnersList(filtered);
    });

    document.getElementById('partners-grid-btn')?.addEventListener('click', () => {
      document.getElementById('partners-grid-btn').classList.add('active');
      document.getElementById('partners-list-btn').classList.remove('active');
      renderPartnersGrid(partnersCache || []);
    });
    document.getElementById('partners-list-btn')?.addEventListener('click', () => {
      document.getElementById('partners-list-btn').classList.add('active');
      document.getElementById('partners-grid-btn').classList.remove('active');
      renderPartnersList(partnersCache || []);
    });

    if (!data.partners || data.partners.length === 0) {
      document.getElementById('partners-container').innerHTML = `
        <div class="dog-state">${dogEmpty('No retail partners yet.')}
          <button class="btn btn-primary" style="margin-top:1rem" onclick="openAddPartnerModal()">✚ Add Partner</button>
        </div>`;
      return;
    }
    renderPartnersGrid(data.partners);
  } catch (e) {
    document.getElementById('partners-container').innerHTML = dogError(e.message);
  }
}

function renderPartnersGrid(partners) {
  const container = document.getElementById('partners-container');
  if (!container) return;
  if (!partners || partners.length === 0) { container.innerHTML = dogEmpty('No partners found.'); return; }
  container.innerHTML = `<div class="partners-grid">${partners.map(p => `
    <div class="partner-card" onclick="openPartnerDetail('${p.value}')">
      <div class="partner-card-body">
        <div class="partner-card-name">${p.label}</div>
        <div class="partner-card-city">${p.city || 'Location not set'}</div>
        <span class="partner-card-split">${getPartnerSplitLabel(p)}</span>
      </div>
    </div>`).join('')}</div>`;
}

function renderPartnersList(partners) {
  const container = document.getElementById('partners-container');
  if (!container) return;
  if (!partners || partners.length === 0) { container.innerHTML = dogEmpty('No partners found.'); return; }
  container.innerHTML = `<div class="card" style="padding:0;overflow:hidden;">
    <table class="items-table">
      <thead><tr><th>Store</th><th>City</th><th>Type</th><th>Revenue Split</th><th>Contact</th></tr></thead>
      <tbody>${partners.map(p => `
        <tr onclick="openPartnerDetail('${p.value}')" style="cursor:pointer;">
          <td><strong>${p.label}</strong></td>
          <td style="color:var(--brown-light)">${p.city || '—'}</td>
          <td><span class="badge badge-teal">${p.partnerType || p.locationType || 'Consignment'}</span></td>
          <td>${getPartnerSplitLabel(p)}</td>
          <td style="color:var(--brown-light)">${p.contactName || '—'}</td>
        </tr>`).join('')}
      </tbody>
    </table>
  </div>`;
}

function getPartnerSplitLabel(p) {
  if (!p) return 'Not set';
  const type = (p.channelType || p.locationType || '').toLowerCase();
  if (type === 'wholesale') return '<span class="partner-card-split">Wholesale</span>';
  if (type === 'market' || type === 'direct') return '<span class="partner-card-split">Direct / Market</span>';
  if (p.split) return '<span class="partner-card-split">' + Math.round(p.split * 100) + '% your cut</span>';
  return '<span class="partner-card-split" style="color:var(--brown-light)">Split not set</span>';
}

function openPartnerDetail(partnerId) {
  const partner = (partnersCache || []).find(p => p.value === partnerId);
  if (!partner) return;
  openDetailPanel(`
    <div class="detail-id">Retail Partner</div>
    <div class="detail-name">${partner.label}</div>
    ${partner.storePhoto ? `<div class="detail-image"><img src="${partner.storePhoto}" alt="${partner.label}"></div>` : ''}
    <div class="detail-stats">
      <div class="detail-stat"><div class="detail-stat-label">Your Cut</div><div class="detail-stat-value">${partner.split ? Math.round(partner.split * 100) + '%' : '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">City</div><div class="detail-stat-value" style="font-size:1rem">${partner.city || '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Contact</div><div class="detail-stat-value" style="font-size:0.85rem">${partner.contactName || '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Phone</div><div class="detail-stat-value" style="font-size:0.85rem">${partner.contactPhone || partner.phone || '—'}</div></div>
    </div>
    ${partner.address ? `<p style="font-size:0.875rem;color:var(--brown-mid);margin-bottom:1rem;">📍 ${partner.address}</p>` : ''}
    ${partner.notes ? `<div style="font-size:0.875rem;color:var(--brown-mid);background:var(--cream);padding:0.75rem;border-radius:var(--radius-sm);margin-bottom:1rem;">📝 ${partner.notes}</div>` : ''}
    ${(() => { try { const codes = JSON.parse(partner.retailItemCodes || '[]'); return codes.length ? `<div style="background:var(--cream);padding:0.75rem;border-radius:var(--radius-sm);margin-bottom:1rem;"><div style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--brown-mid);font-weight:700;margin-bottom:0.4rem;">🏷️ Retail Item Codes</div>${codes.map(c => `<div style="display:flex;justify-content:space-between;font-size:0.85rem;padding:0.2rem 0;"><span style="color:var(--brown-dark)">${c.product}</span><span style="font-family:monospace;color:var(--teal);font-weight:600">${c.code}</span></div>`).join('')}</div>` : ''; } catch(e) { return ''; } })()}
    <hr class="detail-divider">
    <div class="detail-actions">
      <button class="btn btn-primary" onclick="openEditPartnerModal('${partnerId}')">✏️ Edit Partner</button>
      <button class="btn btn-secondary" onclick="closeDetailPanel();document.querySelector('[data-page=record-sale]').click()">◈ Update Inventory</button>
      <button class="btn btn-secondary" onclick="closeDetailPanel()">Close</button>
    </div>`);
}

async function openEditPartnerModal(partnerId) {
  const partner = (partnersCache || []).find(p => p.value === partnerId);
  if (!partner) return;
  let ownerEmail = '', ownerPhone = '';
  try {
    const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getRetailerAuth&locationId=${partnerId}`);
    const d = await r.json();
    if (d.success) { ownerEmail = d.ownerEmail || ''; ownerPhone = d.ownerPhone || ''; }
  } catch (e) {}
  openModal(`
    <div class="modal-title">✏️ Edit Partner</div>
    <form id="edit-partner-form">
      <div class="form-field"><label class="field-label">Store Name *</label><input class="field-input" type="text" name="storeName" value="${partner.label}" required></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">City</label><input class="field-input" type="text" name="city" value="${partner.city || ''}"></div>
        <div class="form-field"><label class="field-label">Your Revenue Split %</label><input class="field-input" type="number" name="split" value="${partner.split ? Math.round(partner.split * 100) : ''}" min="1" max="100"></div>
      </div>
      <div class="form-field">
        <label class="field-label">🏷️ Retail Item Codes</label>
        <div id="sku-codes-list"></div>
        <button type="button" class="btn btn-secondary btn-sm" style="margin-top:0.4rem" onclick="addSkuCodeRow()">+ Add Item Code</button>
      </div>
      <div class="form-field"><label class="field-label">Address</label><input class="field-input" type="text" name="address" value="${partner.address || ''}"></div>
      <div class="form-field"><label class="field-label">Contact Name</label><input class="field-input" type="text" name="contactName" value="${partner.contactName || ''}"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Contact Email</label><input class="field-input" type="email" name="contactEmail" value="${partner.contactEmail || ''}"></div>
        <div class="form-field"><label class="field-label">Store Phone</label><input class="field-input" type="tel" name="contactPhone" value="${partner.contactPhone || partner.phone || ''}" oninput="formatPhoneField(this)"></div>
      </div>
      <div class="form-field"><label class="field-label">Owner Email <span style="color:var(--teal)">(Private — for order verification)</span></label><input class="field-input" type="email" name="ownerEmail" value="${ownerEmail}" placeholder="Never shown to public"></div>
      <div class="form-field"><label class="field-label">Owner Cell <span style="color:var(--teal)">(Private — for order verification)</span></label><input class="field-input" type="tel" name="ownerPhone" value="${ownerPhone}" placeholder="Never shown to public" oninput="formatPhoneField(this)"></div>
      <div class="form-field"><label class="field-label">Notes</label><input class="field-input" type="text" name="notes" value="${partner.notes || ''}"></div>
      <div id="edit-partner-btn-wrap"><button type="submit" class="btn btn-primary" style="width:100%;margin-top:0.5rem">Save Changes</button></div>
      <div id="edit-partner-status" class="form-status"></div>
    </form>`);

  // Populate existing SKU codes
  const existingCodes = (() => { try { return JSON.parse(partner.retailItemCodes || '[]'); } catch(e) { return []; } })();
  if (existingCodes.length === 0) addSkuCodeRow();
  else existingCodes.forEach(c => addSkuCodeRow(c.product, c.code));

  document.getElementById('edit-partner-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status  = document.getElementById('edit-partner-status');
    const btnWrap = document.getElementById('edit-partner-btn-wrap');
    const fd      = new FormData(e.target);
    // Collect SKU codes
    const skuRows = document.querySelectorAll('.sku-code-row');
    const retailItemCodes = [];
    skuRows.forEach(row => {
      const product = row.querySelector('.sku-product')?.value?.trim();
      const code = row.querySelector('.sku-code')?.value?.trim();
      if (product && code) retailItemCodes.push({ product, code });
    });
    if (btnWrap) btnWrap.style.display = 'none';
    status.className  = 'form-status loading'; status.textContent = 'Saving...';
    try {
      const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'updateRetailPartner', partnerData: { locationId: partnerId, storeName: fd.get('storeName'), city: fd.get('city'), split: fd.get('split'), address: fd.get('address'), contactName: fd.get('contactName'), contactEmail: fd.get('contactEmail'), contactPhone: fd.get('contactPhone'), ownerEmail: fd.get('ownerEmail'), ownerPhone: fd.get('ownerPhone'), notes: fd.get('notes'), retailItemCodes: JSON.stringify(retailItemCodes) } }) });
      const result = await r.json();
      if (result.success) {
        partnersCache = null;
        status.className  = 'form-status success'; status.textContent = '✅ Partner updated!';
        showToast('Partner updated!', 'success');
        setTimeout(() => { closeModal(); closeDetailPanel(); renderRetailPartnersPage(); }, 1200);
      } else throw new Error(result.error);
    } catch (err) { status.className = 'form-status error'; status.textContent = '❌ ' + err.message; if (btnWrap) btnWrap.style.display = ''; }
  });
}

function openAddPartnerModal() {
  openModal(`
    <div class="modal-title">✚ Add Retail Partner</div>
    <form id="add-partner-form">
      <div class="form-field"><label class="field-label">Store Name *</label><input class="field-input" type="text" name="storeName" placeholder="e.g. Shenandoah Valley Art Center" required></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">City *</label><input class="field-input" type="text" name="city" placeholder="e.g. Staunton" required></div>
        <div class="form-field"><label class="field-label">Your Revenue Split %</label><input class="field-input" type="number" name="split" placeholder="e.g. 60" min="1" max="100"></div>
      </div>
      <div class="form-field"><label class="field-label">Address</label><input class="field-input" type="text" name="address" placeholder="Street address"></div>
      <div class="form-field"><label class="field-label">Contact Name</label><input class="field-input" type="text" name="contactName" placeholder="Primary contact"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Contact Email</label><input class="field-input" type="email" name="contactEmail" placeholder="email@store.com"></div>
        <div class="form-field"><label class="field-label">Store Phone</label><input class="field-input" type="tel" name="contactPhone" placeholder="555-444-1111" oninput="formatPhoneField(this)"></div>
      </div>
      <div class="form-field"><label class="field-label">Owner Email <span style="color:var(--teal);font-size:0.75rem">(Private — for order verification)</span></label><input class="field-input" type="email" name="ownerEmail" placeholder="Their personal email"></div>
      <div class="form-field"><label class="field-label">Owner Cell <span style="color:var(--teal);font-size:0.75rem">(Private — for order verification)</span></label><input class="field-input" type="tel" name="ownerPhone" placeholder="555-444-1111" oninput="formatPhoneField(this)"></div>
      <div class="form-field"><label class="field-label">Notes</label><input class="field-input" type="text" name="notes" placeholder="Any additional notes"></div>
      <div class="form-field">
        <label class="field-label">🏷️ Retail Item Codes</label>
        <div id="sku-codes-list"></div>
        <button type="button" class="btn btn-secondary btn-sm" style="margin-top:0.4rem" onclick="addSkuCodeRow()">+ Add Item Code</button>
      </div>
      <div id="add-partner-btn-wrap"><button type="submit" class="btn btn-primary" style="width:100%;margin-top:0.5rem">Save Partner</button></div>
      <div id="partner-form-status" class="form-status"></div>
    </form>`);

  addSkuCodeRow(); // Start with one empty row

  document.getElementById('add-partner-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status  = document.getElementById('partner-form-status');
    const btnWrap = document.getElementById('add-partner-btn-wrap');
    const fd      = new FormData(e.target);
    // Collect SKU codes
    const skuRows = document.querySelectorAll('.sku-code-row');
    const retailItemCodes = [];
    skuRows.forEach(row => {
      const product = row.querySelector('.sku-product')?.value?.trim();
      const code = row.querySelector('.sku-code')?.value?.trim();
      if (product && code) retailItemCodes.push({ product, code });
    });
    if (btnWrap) btnWrap.style.display = 'none';
    status.className  = 'form-status loading'; status.textContent = 'Saving...';
    try {
      const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'addRetailPartner', partnerData: { storeName: fd.get('storeName'), city: fd.get('city'), split: fd.get('split'), address: fd.get('address'), contactName: fd.get('contactName'), contactEmail: fd.get('contactEmail'), contactPhone: fd.get('contactPhone'), ownerEmail: fd.get('ownerEmail'), ownerPhone: fd.get('ownerPhone'), notes: fd.get('notes'), retailItemCodes: JSON.stringify(retailItemCodes) } }) });
      const result = await r.json();
      if (result.success) {
        partnersCache = null;
        status.className  = 'form-status success'; status.textContent = '✅ Partner added!';
        showToast('Partner added!', 'success');
        setTimeout(() => { closeModal(); renderRetailPartnersPage(); }, 1200);
      } else throw new Error(result.error);
    } catch (err) { status.className = 'form-status error'; status.textContent = '❌ ' + err.message; if (btnWrap) btnWrap.style.display = ''; }
  });
}

// ============================================================
// VENDING MACHINES
// ============================================================
async function renderVendingMachinesPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Vending Machines</h1><p class="page-subtitle">Track your machines, venues, and removal dates</p></div>
      <div class="page-actions"><button class="btn btn-primary" onclick="openAddMachineModal()">✚ Add Machine</button></div>
    </div>
    <div id="machines-summary" class="machines-summary-bar"></div>
    <div class="items-toolbar" style="margin-bottom:1.25rem;max-width:420px;">
      <div class="search-box" style="flex:1;"><span class="search-icon">⌕</span><input type="text" id="machines-search" placeholder="Search machines..."></div>
    </div>
    <div id="machines-container">${dogLoading('Loading machines...')}</div>`;

  try {
    const r    = await fetch(`${GOOGLE_SCRIPT_URL}?action=getVendingMachines`);
    const data = await r.json();
    if (!data.success) throw new Error(data.error);
    machinesCache = data.machines;
    renderMachinesGrid(data.machines);
    renderMachinesSummary(data.machines);
    document.getElementById('machines-search')?.addEventListener('input', (e) => {
      const q        = e.target.value.toLowerCase();
      const filtered = (machinesCache || []).filter(m => (m.MachineName || '').toLowerCase().includes(q) || (m.VenueName || '').toLowerCase().includes(q) || (m.VenueCity || '').toLowerCase().includes(q));
      renderMachinesGrid(filtered);
    });
  } catch (e) {
    document.getElementById('machines-container').innerHTML = dogError(e.message);
  }
}

function renderMachinesSummary(machines) {
  const bar = document.getElementById('machines-summary');
  if (!bar) return;
  const deployed  = machines.filter(m => m.Status === 'Deployed' || m.Status === 'Active');
  const inStorage = machines.filter(m => m.Status === 'In Storage');
  const upcoming  = machines.filter(m => {
    if (!m.RemovalDate || m.RemovedDate) return false;
    const days = Math.round((new Date(m.RemovalDate) - new Date()) / 86400000);
    return days >= 0 && days <= 14;
  });
  bar.innerHTML = `
    <div class="machines-stats">
      <div class="machine-stat-pill teal"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="16" height="16" style="vertical-align:middle;margin-right:2px"><rect x="12" y="6" width="40" height="52" rx="3" fill="#e03030" stroke="#2d2d2d" stroke-width="2"/><rect x="16" y="10" width="24" height="34" rx="2" fill="#f5f0e8"/><line x1="16" y1="21" x2="40" y2="21" stroke="#ddd" stroke-width="1"/><line x1="16" y1="32" x2="40" y2="32" stroke="#ddd" stroke-width="1"/><circle cx="21" cy="16" r="2.5" fill="#5bc0de"/><circle cx="28" cy="16" r="2.5" fill="#5bc0de"/><circle cx="35" cy="16" r="2.5" fill="#f0c040"/><circle cx="21" cy="27" r="2.5" fill="#e05050"/><circle cx="28" cy="27" r="2.5" fill="#50b080"/><circle cx="35" cy="27" r="2.5" fill="#e05050"/><circle cx="21" cy="38" r="2.5" fill="#f0c040"/><circle cx="28" cy="38" r="2.5" fill="#5bc0de"/><circle cx="35" cy="38" r="2.5" fill="#50b080"/><rect x="42" y="12" width="7" height="16" rx="1.5" fill="#1a1a1a"/><rect x="16" y="48" width="24" height="6" rx="1" fill="#b02020"/></svg> ${deployed.length} Deployed</div>
      <div class="machine-stat-pill amber">📦 ${inStorage.length} In Storage</div>
      ${upcoming.length > 0 ? `<div class="machine-stat-pill coral">⏰ ${upcoming.length} Removal${upcoming.length > 1 ? 's' : ''} Soon</div>` : ''}
    </div>`;
}

function renderMachinesGrid(machines) {
  const container = document.getElementById('machines-container');
  if (!machines || machines.length === 0) {
    container.innerHTML = `<div class="dog-state">${dogEmpty('No vending machines yet.')}<button class="btn btn-primary" style="margin-top:1rem" onclick="openAddMachineModal()">✚ Add Machine</button></div>`;
    return;
  }
  container.innerHTML = `<div class="partners-grid">${machines.map(m => {
    const removalDays = m.RemovalDate && !m.RemovedDate ? Math.round((new Date(m.RemovalDate) - new Date()) / 86400000) : null;
    let removalBadge  = '';
    if (removalDays !== null) {
      if (removalDays < 0)      removalBadge = `<span class="badge badge-coral">Overdue removal</span>`;
      else if (removalDays === 0) removalBadge = `<span class="badge badge-coral">🚨 Remove TODAY</span>`;
      else if (removalDays <= 7)  removalBadge = `<span class="badge badge-amber">⏰ Remove in ${removalDays}d</span>`;
    }
    const statusColor = m.Status === 'Deployed' || m.Status === 'Active' ? 'badge-green' : m.Status === 'In Storage' ? 'badge-teal' : 'badge-amber';
    return `
      <div class="partner-card" onclick="openMachineDetail('${m.MachineID}')">
        <div class="partner-card-image" style="background:var(--cream-dark);display:flex;align-items:center;justify-content:center;">${m.Photo ? `<img src="${fixPhotoUrl(m.Photo)}" alt="${m.MachineName}" style="width:100%;height:100%;object-fit:cover;">` : `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="48" height="48"><rect x="12" y="6" width="40" height="52" rx="3" fill="#e03030" stroke="#2d2d2d" stroke-width="2"/><rect x="16" y="10" width="24" height="34" rx="2" fill="#f5f0e8"/><rect x="16" y="10" width="24" height="34" rx="2" fill="none" stroke="#2d2d2d" stroke-width="1.5"/><line x1="16" y1="21" x2="40" y2="21" stroke="#ddd" stroke-width="1"/><line x1="16" y1="32" x2="40" y2="32" stroke="#ddd" stroke-width="1"/><circle cx="21" cy="16" r="2.5" fill="#5bc0de"/><circle cx="28" cy="16" r="2.5" fill="#5bc0de"/><circle cx="35" cy="16" r="2.5" fill="#f0c040"/><circle cx="21" cy="27" r="2.5" fill="#e05050"/><circle cx="28" cy="27" r="2.5" fill="#50b080"/><circle cx="35" cy="27" r="2.5" fill="#e05050"/><circle cx="21" cy="38" r="2.5" fill="#f0c040"/><circle cx="28" cy="38" r="2.5" fill="#5bc0de"/><circle cx="35" cy="38" r="2.5" fill="#50b080"/><rect x="42" y="12" width="7" height="16" rx="1.5" fill="#1a1a1a"/><rect x="16" y="48" width="24" height="6" rx="1" fill="#b02020"/><rect x="16" y="48" width="24" height="6" rx="1" fill="none" stroke="#2d2d2d" stroke-width="1"/></svg>`}</div>
        <div class="partner-card-body">
          <div class="partner-card-name">${m.MachineName}</div>
          <div class="partner-card-city">${m.Status === 'In Storage' ? '📦 In Storage' : (m.VenueName || 'Venue not set') + (m.VenueCity ? ` — ${m.VenueCity}` : '')}</div>
          <div style="display:flex;gap:0.4rem;flex-wrap:wrap;margin-top:0.5rem;">
            <span class="badge ${statusColor}">${m.Status || 'Unknown'}</span>
            ${removalBadge}
          </div>
          ${m.RemovalDate && !m.RemovedDate ? `<div style="font-size:0.75rem;color:var(--brown-light);margin-top:0.4rem;">Removal: ${new Date(m.RemovalDate).toLocaleDateString()}</div>` : ''}
        </div>
      </div>`;
  }).join('')}</div>`;
}

function openMachineDetail(machineId) {
  const m = (machinesCache || []).find(m => m.MachineID === machineId);
  if (!m) return;
  const removalDays  = m.RemovalDate && !m.RemovedDate ? Math.round((new Date(m.RemovalDate) - new Date()) / 86400000) : null;
  const removalAlert = removalDays !== null && removalDays <= 7
    ? `<div class="removal-alert ${removalDays === 0 ? 'danger' : 'warning'}">${removalDays === 0 ? '🚨 This machine needs to be removed TODAY!' : `⏰ Removal in ${removalDays} day${removalDays === 1 ? '' : 's'}`}</div>` : '';

  openDetailPanel(`
    <div class="detail-id">Vending Machine · ${m.MachineID}</div>
    <div class="detail-name">${m.MachineName}</div>
    <div style="text-align:center;margin:1rem 0;">${m.Photo ? `<img src="${fixPhotoUrl(m.Photo)}" alt="${m.MachineName}" style="max-width:100%;max-height:240px;border-radius:var(--radius-md);object-fit:cover;">` : `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="80" height="80"><rect x="12" y="6" width="40" height="52" rx="3" fill="#e03030" stroke="#2d2d2d" stroke-width="2"/><rect x="16" y="10" width="24" height="34" rx="2" fill="#f5f0e8"/><rect x="16" y="10" width="24" height="34" rx="2" fill="none" stroke="#2d2d2d" stroke-width="1.5"/><line x1="16" y1="21" x2="40" y2="21" stroke="#ddd" stroke-width="1"/><line x1="16" y1="32" x2="40" y2="32" stroke="#ddd" stroke-width="1"/><circle cx="21" cy="16" r="2.5" fill="#5bc0de"/><circle cx="28" cy="16" r="2.5" fill="#5bc0de"/><circle cx="35" cy="16" r="2.5" fill="#f0c040"/><circle cx="21" cy="27" r="2.5" fill="#e05050"/><circle cx="28" cy="27" r="2.5" fill="#50b080"/><circle cx="35" cy="27" r="2.5" fill="#e05050"/><circle cx="21" cy="38" r="2.5" fill="#f0c040"/><circle cx="28" cy="38" r="2.5" fill="#5bc0de"/><circle cx="35" cy="38" r="2.5" fill="#50b080"/><rect x="42" y="12" width="7" height="16" rx="1.5" fill="#1a1a1a"/><rect x="16" y="48" width="24" height="6" rx="1" fill="#b02020"/><rect x="16" y="48" width="24" height="6" rx="1" fill="none" stroke="#2d2d2d" stroke-width="1"/></svg>`}</div>
    ${removalAlert}
    <div class="detail-stats">
      <div class="detail-stat"><div class="detail-stat-label">Status</div><div class="detail-stat-value" style="font-size:1rem">${m.Status || '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Machine Type</div><div class="detail-stat-value" style="font-size:1rem">${m.MachineType || '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Install Date</div><div class="detail-stat-value" style="font-size:0.9rem">${m.InstallDate ? new Date(m.InstallDate).toLocaleDateString() : '—'}</div></div>
      <div class="detail-stat"><div class="detail-stat-label">Removal Date</div><div class="detail-stat-value" style="font-size:0.9rem">${m.RemovalDate ? new Date(m.RemovalDate).toLocaleDateString() : '—'}</div></div>
    </div>
    ${m.VenueName ? `<div class="detail-venue-section"><div class="detail-venue-label">Venue</div><div class="detail-venue-name">${m.VenueName}</div>${m.VenueType ? `<div class="detail-venue-sub">${m.VenueType}</div>` : ''}${m.VenueAddress ? `<div class="detail-venue-sub">📍 ${m.VenueAddress}${m.VenueCity ? ', ' + m.VenueCity : ''}</div>` : ''}${m.ContactName ? `<div class="detail-venue-sub">👤 ${m.ContactName}</div>` : ''}${m.ContactPhone ? `<div class="detail-venue-sub">📞 ${m.ContactPhone}</div>` : ''}${m.ContactEmail ? `<div class="detail-venue-sub">✉️ ${m.ContactEmail}</div>` : ''}</div>` : ''}
    ${m.ResidencyNotes ? `<div style="font-size:0.875rem;color:var(--brown-mid);background:var(--cream);padding:0.75rem;border-radius:var(--radius-sm);margin:1rem 0;">📝 ${m.ResidencyNotes}</div>` : ''}
    <hr class="detail-divider">
    <div class="detail-actions">
      <button class="btn btn-primary" onclick="openEditMachineModal('${machineId}')">✏️ Edit Machine</button>
      ${m.Status !== 'In Storage' && !m.RemovedDate ? `<button class="btn btn-amber" onclick="markMachineRemoved('${machineId}')">📦 Mark as Removed</button>` : ''}
      <button class="btn btn-secondary" onclick="closeDetailPanel()">Close</button>
    </div>`);
}

async function markMachineRemoved(machineId) {
  if (!confirm('Mark this machine as removed and returned to storage?')) return;
  try {
    const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'updateVendingMachine', machineData: { machineId, status: 'In Storage', removedDate: new Date().toLocaleDateString('en-CA') } }) });
    const result = await r.json();
    if (result.success) { machinesCache = null; showToast('Machine marked as removed!', 'success'); closeDetailPanel(); renderVendingMachinesPage(); }
    else throw new Error(result.error);
  } catch (e) { showToast('Error: ' + e.message, 'error'); }
}

function openAddMachineModal() {
  openModal(`
    <div class="modal-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="22" height="22" style="vertical-align:middle;margin-right:4px"><rect x="12" y="6" width="40" height="52" rx="3" fill="#e03030" stroke="#2d2d2d" stroke-width="2"/><rect x="16" y="10" width="24" height="34" rx="2" fill="#f5f0e8"/><line x1="16" y1="21" x2="40" y2="21" stroke="#ddd" stroke-width="1"/><line x1="16" y1="32" x2="40" y2="32" stroke="#ddd" stroke-width="1"/><circle cx="21" cy="16" r="2.5" fill="#5bc0de"/><circle cx="28" cy="16" r="2.5" fill="#5bc0de"/><circle cx="35" cy="16" r="2.5" fill="#f0c040"/><circle cx="21" cy="27" r="2.5" fill="#e05050"/><circle cx="28" cy="27" r="2.5" fill="#50b080"/><circle cx="35" cy="27" r="2.5" fill="#e05050"/><circle cx="21" cy="38" r="2.5" fill="#f0c040"/><circle cx="28" cy="38" r="2.5" fill="#5bc0de"/><circle cx="35" cy="38" r="2.5" fill="#50b080"/><rect x="42" y="12" width="7" height="16" rx="1.5" fill="#1a1a1a"/><rect x="16" y="48" width="24" height="6" rx="1" fill="#b02020"/></svg> Add Vending Machine</div>
    <form id="add-machine-form">
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Machine Name *</label><input class="field-input" type="text" name="machineName" placeholder="e.g. Sasha" required></div>
        <div class="form-field"><label class="field-label">Machine Type</label><input class="field-input" type="text" name="machineType" placeholder="e.g. Standard, Mini"></div>
      </div>
      <div class="form-field"><label class="field-label">Status</label><select class="field-input" name="status"><option value="In Storage">In Storage</option><option value="Deployed">Deployed</option></select></div>
      <div class="form-field"><label class="field-label">Venue Name</label><input class="field-input" type="text" name="venueName" placeholder="e.g. The Artisan Market"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Venue Type</label><input class="field-input" type="text" name="venueType" placeholder="e.g. Art Gallery"></div>
        <div class="form-field"><label class="field-label">City</label><input class="field-input" type="text" name="venueCity" placeholder="e.g. Staunton"></div>
      </div>
      <div class="form-field"><label class="field-label">Address</label><input class="field-input" type="text" name="venueAddress" placeholder="Street address"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Contact Name</label><input class="field-input" type="text" name="contactName"></div>
        <div class="form-field"><label class="field-label">Contact Phone</label><input class="field-input" type="tel" name="contactPhone" oninput="formatPhoneField(this)"></div>
      </div>
      <div class="form-field"><label class="field-label">Contact Email</label><input class="field-input" type="email" name="contactEmail"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Install Date</label><input class="field-input" type="date" name="installDate"></div>
        <div class="form-field"><label class="field-label">Removal Date</label><input class="field-input" type="date" name="removalDate"></div>
      </div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Notify Email</label><input class="field-input" type="email" name="notifyEmail"></div>
        <div class="form-field"><label class="field-label">Notify Phone (SMS)</label><input class="field-input" type="tel" name="notifyPhone" oninput="formatPhoneField(this)"></div>
      </div>
      <div class="form-field"><label class="field-label">Residency Notes</label><input class="field-input" type="text" name="residencyNotes"></div>
      <div class="form-field">
        <label class="field-label">Photo</label>
        <div class="photo-upload-area" onclick="document.getElementById('machine-photo-file').click()">
          <input type="file" id="machine-photo-file" accept="image/*" style="display:none" onchange="handleMachinePhotoUpload(this)">
          <div class="photo-upload-icon">📷</div>
          <div class="photo-upload-text">Tap to take photo or choose from device</div>
        </div>
        <input class="field-input" type="text" name="photo" id="machine-photo-url" placeholder="Or paste a URL..." style="margin-top:0.5rem;" onchange="handlePhotoUrlPaste(this, '.machine-photo-preview-wrap')" onpaste="setTimeout(() => handlePhotoUrlPaste(this, '.machine-photo-preview-wrap'), 100)">
        <div id="machine-photo-status" style="font-size:0.75rem;margin-top:0.3rem;display:none;"></div>
      </div>
      <div id="add-machine-btn-wrap"><button type="submit" class="btn btn-primary" style="width:100%;margin-top:0.5rem">Save Machine</button></div>
      <div id="machine-form-status" class="form-status"></div>
    </form>`);

  document.getElementById('add-machine-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status  = document.getElementById('machine-form-status');
    const btnWrap = document.getElementById('add-machine-btn-wrap');
    const fd      = new FormData(e.target);
    if (btnWrap) btnWrap.style.display = 'none';
    status.className  = 'form-status loading'; status.textContent = 'Saving...';
    try {
      const machineData = Object.fromEntries(fd.entries()); if (machineData.photo) machineData.photo = fixPhotoUrl(machineData.photo);
      const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'addVendingMachine', machineData }) });
      const result = await r.json();
      if (result.success) { machinesCache = null; status.className = 'form-status success'; status.textContent = `✅ ${fd.get('machineName')} added!`; showToast('Machine added!', 'success'); setTimeout(() => { closeModal(); renderVendingMachinesPage(); }, 1200); }
      else throw new Error(result.error);
    } catch (err) { status.className = 'form-status error'; status.textContent = '❌ ' + err.message; if (btnWrap) btnWrap.style.display = ''; }
  });
}

function openEditMachineModal(machineId) {
  const m   = (machinesCache || []).find(m => m.MachineID === machineId);
  if (!m) return;
  const fmt = (d) => d ? new Date(d).toLocaleDateString('en-CA') : '';
  openModal(`
    <div class="modal-title">✏️ Edit ${m.MachineName}</div>
    <form id="edit-machine-form">
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Machine Name *</label><input class="field-input" type="text" name="machineName" value="${m.MachineName || ''}" required></div>
        <div class="form-field"><label class="field-label">Machine Type</label><input class="field-input" type="text" name="machineType" value="${m.MachineType || ''}"></div>
      </div>
      <div class="form-field"><label class="field-label">Status</label><select class="field-input" name="status"><option value="In Storage" ${m.Status === 'In Storage' ? 'selected' : ''}>In Storage</option><option value="Deployed" ${m.Status === 'Deployed' ? 'selected' : ''}>Deployed</option><option value="Active" ${m.Status === 'Active' ? 'selected' : ''}>Active</option></select></div>
      <div class="form-field"><label class="field-label">Venue Name</label><input class="field-input" type="text" name="venueName" value="${m.VenueName || ''}"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Venue Type</label><input class="field-input" type="text" name="venueType" value="${m.VenueType || ''}"></div>
        <div class="form-field"><label class="field-label">City</label><input class="field-input" type="text" name="venueCity" value="${m.VenueCity || ''}"></div>
      </div>
      <div class="form-field"><label class="field-label">Address</label><input class="field-input" type="text" name="venueAddress" value="${m.VenueAddress || ''}"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Contact Name</label><input class="field-input" type="text" name="contactName" value="${m.ContactName || ''}"></div>
        <div class="form-field"><label class="field-label">Contact Phone</label><input class="field-input" type="tel" name="contactPhone" value="${m.ContactPhone || ''}" oninput="formatPhoneField(this)"></div>
      </div>
      <div class="form-field"><label class="field-label">Contact Email</label><input class="field-input" type="email" name="contactEmail" value="${m.ContactEmail || ''}"></div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Install Date</label><input class="field-input" type="date" name="installDate" value="${fmt(m.InstallDate)}"></div>
        <div class="form-field"><label class="field-label">Removal Date</label><input class="field-input" type="date" name="removalDate" value="${fmt(m.RemovalDate)}"></div>
      </div>
      <div class="form-grid">
        <div class="form-field"><label class="field-label">Notify Email</label><input class="field-input" type="email" name="notifyEmail" value="${m.NotifyEmail || ''}"></div>
        <div class="form-field"><label class="field-label">Notify Phone</label><input class="field-input" type="tel" name="notifyPhone" value="${m.NotifyPhone || ''}" oninput="formatPhoneField(this)"></div>
      </div>
      <div class="form-field"><label class="field-label">Residency Notes</label><input class="field-input" type="text" name="residencyNotes" value="${m.ResidencyNotes || ''}"></div>
      <div class="form-field">
        <label class="field-label">Photo</label>
        ${m.Photo ? `<div class="photo-preview-wrap"><img class="photo-preview-thumb" src="${fixPhotoUrl(m.Photo)}" alt="Machine photo"></div>` : ''}
        <div class="photo-upload-area" onclick="document.getElementById('machine-photo-file').click()">
          <input type="file" id="machine-photo-file" accept="image/*" style="display:none" onchange="handleMachinePhotoUpload(this)">
          <div class="photo-upload-icon">📷</div>
          <div class="photo-upload-text">Tap to take photo or choose from device</div>
        </div>
        <input class="field-input" type="text" name="photo" id="machine-photo-url" value="${m.Photo || ''}" placeholder="Or paste a URL..." style="margin-top:0.5rem;" onchange="handlePhotoUrlPaste(this, '.machine-photo-preview-wrap')" onpaste="setTimeout(() => handlePhotoUrlPaste(this, '.machine-photo-preview-wrap'), 100)">
        <div id="machine-photo-status" style="font-size:0.75rem;margin-top:0.3rem;display:none;"></div>
      </div>
      <div id="edit-machine-btn-wrap"><button type="submit" class="btn btn-primary" style="width:100%;margin-top:0.5rem">Save Changes</button></div>
      <div id="edit-machine-status" class="form-status"></div>
    </form>`);

  document.getElementById('edit-machine-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status  = document.getElementById('edit-machine-status');
    const btnWrap = document.getElementById('edit-machine-btn-wrap');
    const fd      = new FormData(e.target);
    if (btnWrap) btnWrap.style.display = 'none';
    status.className  = 'form-status loading'; status.textContent = 'Saving...';
    try {
      const machineUpdate = { machineId, ...Object.fromEntries(fd.entries()) }; if (machineUpdate.photo) machineUpdate.photo = fixPhotoUrl(machineUpdate.photo);
      const r      = await fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'updateVendingMachine', machineData: machineUpdate }) });
      const result = await r.json();
      if (result.success) { machinesCache = null; status.className = 'form-status success'; status.textContent = '✅ Machine updated!'; showToast('Machine updated!', 'success'); setTimeout(() => { closeModal(); closeDetailPanel(); renderVendingMachinesPage(); }, 1200); }
      else throw new Error(result.error);
    } catch (err) { status.className = 'form-status error'; status.textContent = '❌ ' + err.message; if (btnWrap) btnWrap.style.display = ''; }
  });
}

// ============================================================
// ORDERS
// ============================================================
async function renderOrdersPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Orders & Requests</h1><p class="page-subtitle">Review and fulfill retailer stock requests</p></div>
    </div>
    <div style="display:flex;align-items:center;gap:0.75rem;margin-bottom:1.25rem;flex-wrap:wrap;">
      <div class="orders-filter-bar">
        <button class="period-btn active" onclick="filterOrders('all', this)">All</button>
        <button class="period-btn" onclick="filterOrders('Pending', this)">Pending</button>
        <button class="period-btn" onclick="filterOrders('Fulfilled', this)">Fulfilled</button>
      </div>
      <button class="filter-btn" id="orders-filter-btn" onclick="openOrdersFilterPanel()">
        <svg width="15" height="15" viewBox="0 0 15 15" fill="none"><path d="M1 3.5h13M3.5 7.5h8M6 11.5h3" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
        Filter by Partner
        <span class="filter-count-badge" id="orders-filter-count" style="display:none">0</span>
      </button>
    </div>
    <div id="orders-container">${dogLoading('Loading orders...')}</div>

    <div class="filter-panel-overlay" id="orders-filter-overlay" onclick="closeOrdersFilterPanel()"></div>
    <div class="filter-panel" id="orders-filter-panel">
      <div class="filter-panel-header">
        <span class="filter-panel-title">Filter Orders</span>
        <button class="filter-panel-close" onclick="closeOrdersFilterPanel()">✕</button>
      </div>
      <div class="filter-panel-body" id="orders-filter-body">
        <div class="filter-category">
          <div class="filter-category-label">Partner Type</div>
          <div class="filter-tag-chips">
            <button class="filter-tag-chip" onclick="toggleOrdersFilter('type','wholesale',this)">Wholesale</button>
            <button class="filter-tag-chip" onclick="toggleOrdersFilter('type','consignment',this)">Consignment</button>
          </div>
        </div>
        <div class="filter-category" id="orders-partner-list">
          <div class="filter-category-label">By Partner</div>
          <div class="filter-tag-chips" id="orders-partner-chips">Loading...</div>
        </div>
      </div>
      <div class="filter-panel-footer">
        <button class="btn btn-secondary" onclick="clearOrdersFilters()">Clear All</button>
        <button class="btn btn-primary" onclick="applyOrdersFilters()">Show Results</button>
      </div>
    </div>`;

  try {
    const r    = await fetch(`${GOOGLE_SCRIPT_URL}?action=getOrders`);
    const data = await r.json();
    if (!data.success) throw new Error(data.error);
    window._ordersCache = data.orders;
    renderOrdersList(data.orders);
  } catch (e) {
    document.getElementById('orders-container').innerHTML = dogError(e.message);
  }
}

let activeOrdersFilters = { types: new Set(), partners: new Set() };

function openOrdersFilterPanel() {
  const orders   = window._ordersCache || [];
  const partners = [...new Set(orders.map(o => o.PartnerName).filter(Boolean))];
  const chipsEl  = document.getElementById('orders-partner-chips');
  if (chipsEl) {
    chipsEl.innerHTML = partners.length === 0
      ? '<span style="font-size:0.8rem;color:var(--brown-light)">No orders yet</span>'
      : partners.map(p => `<button class="filter-tag-chip ${activeOrdersFilters.partners.has(p) ? 'selected' : ''}" onclick="toggleOrdersFilter('partner','${p.replace(/'/g,"\\'")}',this)">${p}</button>`).join('');
  }
  document.getElementById('orders-filter-panel').classList.add('open');
  document.getElementById('orders-filter-overlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeOrdersFilterPanel() {
  document.getElementById('orders-filter-panel').classList.remove('open');
  document.getElementById('orders-filter-overlay').classList.remove('open');
  document.body.style.overflow = '';
}

function toggleOrdersFilter(key, val, btn) {
  const set = key === 'type' ? activeOrdersFilters.types : activeOrdersFilters.partners;
  if (set.has(val)) { set.delete(val); btn.classList.remove('selected'); }
  else { set.add(val); btn.classList.add('selected'); }
}

function clearOrdersFilters() {
  activeOrdersFilters = { types: new Set(), partners: new Set() };
  document.querySelectorAll('#orders-filter-panel .filter-tag-chip').forEach(c => c.classList.remove('selected'));
}

function applyOrdersFilters() {
  closeOrdersFilterPanel();
  const count = activeOrdersFilters.types.size + activeOrdersFilters.partners.size;
  const badge = document.getElementById('orders-filter-count');
  if (badge) { badge.textContent = count; badge.style.display = count > 0 ? 'inline-flex' : 'none'; }
  document.getElementById('orders-filter-btn')?.classList.toggle('has-filters', count > 0);
  let orders = window._ordersCache || [];
  if (activeOrdersFilters.types.size > 0)    orders = orders.filter(o => activeOrdersFilters.types.has(o.PartnerType));
  if (activeOrdersFilters.partners.size > 0) orders = orders.filter(o => activeOrdersFilters.partners.has(o.PartnerName));
  renderOrdersList(orders);
}

function filterOrders(status, btn) {
  document.querySelectorAll('.orders-filter-bar .period-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  const orders   = window._ordersCache || [];
  const filtered = status === 'all' ? orders : orders.filter(o => o.Status === status);
  renderOrdersList(filtered);
}

function renderOrderCard(o) {
  const isPending    = o.Status === 'Pending';
  const isWholesale  = o.PartnerType === 'wholesale';
  const submittedDate = o.SubmittedAt ? new Date(o.SubmittedAt).toLocaleDateString() : '—';
  return `
    <div class="order-card ${isPending ? 'order-pending' : ''}">
      <div class="order-card-header">
        <div>
          <div class="order-id">${o.OrderID}</div>
          <div class="order-partner">${o.PartnerName || '—'}</div>
          <div class="order-meta">${submittedDate} · ${o.SubmitterName || '—'} (${o.SubmitterEmail || '—'})</div>
        </div>
        <div style="text-align:right;">
          <span class="badge ${isPending ? 'badge-amber' : 'badge-green'}">${o.Status}</span>
          <div style="margin-top:0.3rem;">
            <span class="badge ${isWholesale ? 'badge-teal' : 'badge-coral'}">${isWholesale ? 'Wholesale' : 'Consignment'}</span>
          </div>
        </div>
      </div>
      <div class="order-items-list">
        ${(o.items || []).map(item => `
          <div class="order-item-row">
            <span class="order-item-id">#${item.itemId}</span>
            <span class="order-item-name">${item.designName}</span>
            <span class="order-item-qty">× ${item.qty}</span>
            ${isWholesale ? `<span class="order-item-price">$${((item.wholesalePrice || 2) * item.qty).toFixed(2)}</span>` : ''}
          </div>`).join('')}
      </div>
      ${isWholesale ? `
        <div class="order-totals">
          <div>Subtotal: <strong>$${o.SubTotal || '—'}</strong></div>
          <div>Tax (5.3%): <strong>$${o.TaxAmount || '—'}</strong></div>
          <div>Est. Total: <strong>$${o.EstTotal || '—'}</strong></div>
        </div>` : ''}
      ${isPending ? `
        <div class="order-actions">
          <button class="btn btn-primary btn-sm" onclick="openFulfillModal('${o.OrderID}')">✓ Fulfill Order</button>
        </div>` : `
        <div class="order-fulfilled-note" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;">
          <span>Fulfilled ${o.FulfilledAt ? new Date(o.FulfilledAt).toLocaleDateString() : ''}</span>
          <button class="btn btn-secondary btn-sm" onclick="undoFulfillment('${o.OrderID}')" style="font-size:0.75rem;padding:0.25rem 0.6rem;">↩ Undo</button>
        </div>`}
    </div>`;
}

function renderOrdersList(orders) {
  const container = document.getElementById('orders-container');
  if (!orders || orders.length === 0) {
    container.innerHTML = `<div class="dog-state">${dogEmpty('No orders yet.')}</div>`;
    return;
  }

  const pending   = orders.filter(o => o.Status === 'Pending');
  const fulfilled = orders.filter(o => o.Status !== 'Pending');

  // Group fulfilled by year
  const fulfilledByYear = {};
  fulfilled.forEach(o => {
    const year = o.FulfilledAt ? new Date(o.FulfilledAt).getFullYear() : (o.SubmittedAt ? new Date(o.SubmittedAt).getFullYear() : 'Unknown');
    if (!fulfilledByYear[year]) fulfilledByYear[year] = [];
    fulfilledByYear[year].push(o);
  });
  const sortedYears = Object.keys(fulfilledByYear).sort((a, b) => b - a);

  let html = '';

  // Pending orders section
  if (pending.length > 0) {
    html += `<div class="orders-section-label" style="font-size:0.85rem;font-weight:600;color:var(--amber-dark);margin-bottom:0.75rem;display:flex;align-items:center;gap:0.5rem;">
      <span style="width:8px;height:8px;border-radius:50%;background:var(--amber);display:inline-block;"></span>
      Pending (${pending.length})
    </div>`;
    html += pending.map(o => renderOrderCard(o)).join('');
  } else {
    html += `<div style="text-align:center;padding:1.5rem;color:var(--brown-light);font-size:0.875rem;">No pending orders</div>`;
  }

  // Divider
  if (fulfilled.length > 0) {
    html += `<div style="border-top:2px solid var(--brown-200);margin:1.5rem 0;"></div>`;

    // Fulfilled header — collapsible
    html += `<div class="orders-section-label" style="font-size:0.85rem;font-weight:600;color:var(--green-dark,var(--teal-dark));margin-bottom:0.75rem;cursor:pointer;display:flex;align-items:center;gap:0.5rem;user-select:none;"
      onclick="toggleFulfilledSection()">
      <svg id="fulfilled-chevron" width="14" height="14" viewBox="0 0 14 14" fill="none" style="transition:transform 0.2s;transform:rotate(-90deg);">
        <path d="M4.5 2.5L9.5 7L4.5 11.5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
      Fulfilled (${fulfilled.length})
    </div>`;

    html += `<div id="fulfilled-orders-section" style="display:none;">`;

    sortedYears.forEach(year => {
      const yearOrders = fulfilledByYear[year];
      html += `
        <div class="fulfilled-year-group" style="margin-bottom:1rem;">
          <div style="font-size:0.8rem;font-weight:600;color:var(--brown-mid);cursor:pointer;display:flex;align-items:center;gap:0.4rem;padding:0.4rem 0;user-select:none;"
            onclick="toggleYearGroup('${year}')">
            <svg class="year-chevron" id="year-chevron-${year}" width="12" height="12" viewBox="0 0 14 14" fill="none" style="transition:transform 0.2s;transform:rotate(-90deg);">
              <path d="M4.5 2.5L9.5 7L4.5 11.5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
            ${year} (${yearOrders.length})
          </div>
          <div id="year-group-${year}" style="display:none;">
            ${yearOrders.map(o => renderOrderCard(o)).join('')}
          </div>
        </div>`;
    });

    html += `</div>`;
  }

  container.innerHTML = html;
}

function toggleFulfilledSection() {
  const section = document.getElementById('fulfilled-orders-section');
  const chevron = document.getElementById('fulfilled-chevron');
  if (!section) return;
  const isHidden = section.style.display === 'none';
  section.style.display = isHidden ? 'block' : 'none';
  if (chevron) chevron.style.transform = isHidden ? 'rotate(90deg)' : 'rotate(-90deg)';
}

function toggleYearGroup(year) {
  const group   = document.getElementById('year-group-' + year);
  const chevron = document.getElementById('year-chevron-' + year);
  if (!group) return;
  const isHidden = group.style.display === 'none';
  group.style.display = isHidden ? 'block' : 'none';
  if (chevron) chevron.style.transform = isHidden ? 'rotate(90deg)' : 'rotate(-90deg)';
}

async function undoFulfillment(orderId) {
  if (!confirm('Undo fulfillment for ' + orderId + '? This will set the order back to Pending. (Stock will NOT be automatically restored.)')) return;
  try {
    showToast('Undoing fulfillment...', 'info');
    const r = await fetch(GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'unfulfillOrder', orderId })
    });
    const result = await r.json();
    if (result.success) {
      showToast('Order set back to Pending!', 'success');
      renderOrdersPage();
    } else {
      showToast('Error: ' + result.error, 'error');
    }
  } catch (err) {
    showToast('Error: ' + err.message, 'error');
  }
}

// ============================================================
// FULFILL ORDER MODAL
// ============================================================
function openFulfillModal(orderId) {
  const order = (window._ordersCache || []).find(o => o.OrderID === orderId);
  if (!order) return;

  const isConsignment = order.PartnerType !== 'wholesale';
  const partnerRecord = (partnersCache || []).find(p => p.value === order.LocationID);
  const partnerEmail  = partnerRecord?.contactEmail || '';
  const today         = new Date().toLocaleDateString('en-CA');
  const nextWeek      = new Date(Date.now() + 7 * 86400000).toLocaleDateString('en-CA');

  openModal(`
    <div class="modal-title">✓ Fulfill Order</div>
    <p style="font-size:0.875rem;color:var(--brown-mid);margin-bottom:1.25rem;">
      Review and adjust quantities if needed. Once you click Confirm, stock will be deducted.
    </p>
    <form id="fulfill-form">
      <div id="fulfill-items">
        ${(order.items || []).map(item => `
          <div class="fulfill-item-row">
            <div class="fulfill-item-info">
              <strong>#${item.itemId}</strong> ${item.designName}
            </div>
            <div class="fulfill-item-qty">
              <label style="font-size:0.75rem;color:var(--brown-light)">Actual Qty</label>
              <input type="number" class="field-input" name="qty_${item.itemId}"
                value="${item.qty}" min="0" style="width:80px;text-align:center;">
            </div>
            ${isConsignment ? `
              <div style="display:flex;align-items:center;gap:0.5rem;margin-top:0.25rem;">
                <label style="font-size:0.75rem;color:var(--brown-light);white-space:nowrap;">Status:</label>
                <select class="field-input item-approval-select" name="approval_${item.itemId}"
                  style="font-size:0.8rem;padding:0.3rem 0.5rem;">
                  <option value="approved">✅ Approved</option>
                  <option value="declined">❌ Declined</option>
                </select>
                <input type="text" class="field-input item-decline-reason" name="reason_${item.itemId}"
                  placeholder="Reason (if declined)"
                  style="font-size:0.8rem;display:none;flex:1;">
              </div>` : ''}
          </div>`).join('')}
      </div>

      <div class="form-field" style="margin-top:1rem;">
        <label class="field-label">Fulfillment Notes</label>
        <input class="field-input" type="text" name="fulfillNotes" placeholder="Any internal notes about this fulfillment">
      </div>

      ${isConsignment ? `
        <div style="margin-top:1.25rem;background:var(--teal-light);border:1.5px solid var(--teal);border-radius:var(--radius-md);padding:1.25rem;">
          <div style="font-size:1rem;color:var(--teal-dark);margin-bottom:0.75rem;font-weight:600;">
            📬 Send Restock Approval to Partner?
          </div>
          <label style="display:flex;align-items:center;gap:0.5rem;margin-bottom:1rem;cursor:pointer;font-size:0.875rem;color:var(--brown-mid);">
            <input type="checkbox" id="send-notif-toggle" onchange="toggleNotifSection(this)">
            Yes, send approval email to this partner
          </label>
          <div id="notif-section" style="display:none;">
            <div class="form-field">
              <label class="field-label">Partner Contact Email</label>
              <input class="field-input" type="email" id="notif-partner-email" value="${partnerEmail}" placeholder="partner@store.com">
              <div class="field-hint">Email address the approval will be sent to</div>
            </div>
            <div class="form-grid" style="margin-bottom:1rem;">
              <div class="form-field">
                <label class="field-label">Visit Window — From</label>
                <input class="field-input" type="date" id="notif-date-from" value="${today}">
              </div>
              <div class="form-field">
                <label class="field-label">Visit Window — To</label>
                <input class="field-input" type="date" id="notif-date-to" value="${nextWeek}">
              </div>
            </div>
            <div class="form-field">
              <label class="field-label">Personal Note from Angel (optional)</label>
              <input class="field-input" type="text" id="notif-admin-note"
                placeholder="e.g. Can't wait to see you! Bringing new holiday cards too.">
            </div>
          </div>
        </div>` : ''}

      <div id="fulfill-btn-wrap"><button type="submit" class="btn btn-primary" style="width:100%;margin-top:1rem;">Confirm Fulfillment</button></div>
      <div id="fulfill-status" class="form-status"></div>
    </form>`);

  if (isConsignment) {
    document.querySelectorAll('.item-approval-select').forEach(sel => {
      sel.addEventListener('change', function() {
        const itemId      = this.name.replace('approval_', '');
        const reasonInput = document.querySelector(`[name="reason_${itemId}"]`);
        if (reasonInput) reasonInput.style.display = this.value === 'declined' ? 'block' : 'none';
      });
    });
  }

  document.getElementById('fulfill-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status  = document.getElementById('fulfill-status');
    const btnWrap = document.getElementById('fulfill-btn-wrap');
    const fd      = new FormData(e.target);
    if (btnWrap) btnWrap.style.display = 'none';
    status.className  = 'form-status loading';
    status.textContent = 'Fulfilling...';

    const adjustedItems = (order.items || []).map(item => ({
      ...item,
      qty: parseInt(fd.get(`qty_${item.itemId}`) || item.qty)
    }));

    let notifData = null;
    if (isConsignment && document.getElementById('send-notif-toggle')?.checked) {
      const approvedItems = [];
      const declinedItems = [];
      (order.items || []).forEach(item => {
        const approval = fd.get(`approval_${item.itemId}`);
        const actualQty = parseInt(fd.get(`qty_${item.itemId}`) || item.qty);
        if (approval === 'declined') {
          declinedItems.push({ itemId: item.itemId, designName: item.designName, reason: fd.get(`reason_${item.itemId}`) || 'Out of stock' });
        } else {
          approvedItems.push({ itemId: item.itemId, designName: item.designName, approvedQty: actualQty, notes: '' });
        }
      });
      notifData = {
        partnerName:   order.PartnerName,
        partnerEmail:  document.getElementById('notif-partner-email')?.value.trim(),
        visitDateFrom: document.getElementById('notif-date-from')?.value,
        visitDateTo:   document.getElementById('notif-date-to')?.value,
        approvedItems,
        declinedItems,
        adminNote:     document.getElementById('notif-admin-note')?.value.trim(),
      };
      if (!notifData.partnerEmail) {
        status.className  = 'form-status error';
        status.textContent = '❌ Please enter a partner email address.';
        if (btnWrap) btnWrap.style.display = '';
        return;
      }
    }

    try {
      const r      = await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'fulfillOrder', orderId, adjustments: adjustedItems, notifData })
      });
      const result = await r.json();
      if (result.success) {
        status.className  = 'form-status success';
        status.textContent = notifData ? '✅ Order fulfilled & approval email sent!' : '✅ Order fulfilled! Stock updated.';
        showToast('Order fulfilled!', 'success');
        setTimeout(() => { closeModal(); renderOrdersPage(); }, 1600);
      } else throw new Error(result.error);
    } catch (err) {
      status.className  = 'form-status error';
      status.textContent = '❌ ' + err.message;
      if (btnWrap) btnWrap.style.display = '';
    }
  });
}

function toggleNotifSection(checkbox) {
  const section = document.getElementById('notif-section');
  if (section) section.style.display = checkbox.checked ? 'block' : 'none';
}

// ============================================================
// SALES REPORTS
// ============================================================
// ============================================================
// MARKET SALES
// ============================================================
function renderMarketSalesPage() {
  const now = new Date();
  const todayStr = now.toISOString().split('T')[0];

  appContainer.innerHTML = `
    <div class="page-header">
      <div>
        <h1 class="page-title">Market Sales</h1>
        <p class="page-subtitle">Log sales from art markets, pop-ups, and direct sales</p>
      </div>
    </div>
    <div class="form-page-container">
      <form id="mkt-sale-form">
        <div class="form-section">
          <div class="form-section-title">Log a Sale</div>
          <div class="form-grid">
            <div class="form-field">
              <label class="field-label">Date *</label>
              <input class="field-input" type="date" id="mkt-date" value="${todayStr}" required>
            </div>
            <div class="form-field">
              <label class="field-label">Market / Event Name</label>
              <input class="field-input" type="text" id="mkt-name" placeholder="e.g. Downtown Art Walk">
            </div>
          </div>
          <div class="form-grid">
            <div class="form-field">
              <label class="field-label">Total Sales ($) *</label>
              <input class="field-input" type="number" id="mkt-total" step="0.01" min="0" placeholder="0.00" required>
            </div>
            <div class="form-field">
              <label class="field-label">Misprint Sales ($)</label>
              <input class="field-input" type="number" id="mkt-misprint" step="0.01" min="0" placeholder="0.00" value="0">
            </div>
          </div>
          <div class="form-field">
            <label class="field-label">Cards Sold</label>
            <input class="field-input" type="number" id="mkt-cards" min="0" placeholder="0" style="max-width:160px;">
          </div>
        </div>
        <div id="mkt-btn-wrap"><button type="submit" class="btn btn-primary btn-lg" style="width:100%">🎪 Save Market Sale</button></div>
        <div id="mkt-status" class="form-status"></div>
      </form>

      <div class="form-section" style="margin-top:1.5rem;">
        <div class="form-section-title">Recent Sales</div>
        <div id="mkt-history">${dogLoading('Loading sales history...')}</div>
      </div>
    </div>`;

  document.getElementById('mkt-sale-form').addEventListener('submit', (e) => { e.preventDefault(); submitMarketSale(); });
  loadMarketSalesHistory();
}

async function submitMarketSale() {
  const date      = document.getElementById('mkt-date').value;
  const totalSales = document.getElementById('mkt-total').value;
  const btnWrap   = document.getElementById('mkt-btn-wrap');
  const status    = document.getElementById('mkt-status');

  if (!date || !totalSales) {
    status.textContent = 'Please enter a date and total sales amount.';
    status.style.color = '#c33';
    return;
  }

  btnWrap.style.display = 'none';
  status.textContent = 'Saving...';
  status.style.color = 'var(--brown-light)';

  try {
    const payload = {
      action: 'logMarketSale',
      date,
      marketName: document.getElementById('mkt-name').value.trim(),
      totalSales: parseFloat(totalSales) || 0,
      misprintSales: parseFloat(document.getElementById('mkt-misprint').value) || 0,
      cardsSold: parseInt(document.getElementById('mkt-cards').value) || 0
    };

    const r = await fetch(GOOGLE_SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify(payload)
    });
    const data = await r.json();

    if (data.success) {
      showToast('Market sale logged!');
      status.textContent = '';
      btnWrap.style.display = '';
      document.getElementById('mkt-name').value = '';
      document.getElementById('mkt-total').value = '';
      document.getElementById('mkt-misprint').value = '0';
      document.getElementById('mkt-cards').value = '';
      loadMarketSalesHistory();
    } else {
      status.textContent = 'Error: ' + (data.error || 'Unknown');
      status.style.color = '#c33';
      btnWrap.style.display = '';
    }
  } catch (e) {
    status.textContent = 'Network error. Try again.';
    status.style.color = '#c33';
    btnWrap.style.display = '';
  }
}

async function loadMarketSalesHistory() {
  const container = document.getElementById('mkt-history');
  if (!container) return;

  try {
    const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getMarketSales`);
    const data = await r.json();

    if (!data.success || !data.sales || data.sales.length === 0) {
      container.innerHTML = `<div class="dog-state" style="padding:2rem">${dogEmpty('No market sales logged yet')}</div>`;
      return;
    }

    // Sort newest first
    const sorted = data.sales.sort((a, b) => String(b.Date).localeCompare(String(a.Date)));

    container.innerHTML = sorted.map(s => {
      const dateStr = s.Date ? new Date(s.Date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '—';
      const name = s.MarketName || 'Direct Sale';
      const total = parseFloat(s.TotalSales) || 0;
      const misprint = parseFloat(s.MisprintSales) || 0;
      const cards = parseInt(s.CardsSold) || 0;
      return `
        <div class="item-card" style="padding:0.85rem 1rem;margin-bottom:0.5rem;cursor:default;">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <div>
              <div style="font-weight:600;color:var(--brown-dark)">${name}</div>
              <div style="font-size:0.78rem;color:var(--brown-light);margin-top:2px">${dateStr}${cards ? ' · ' + cards + ' cards' : ''}</div>
            </div>
            <div style="text-align:right">
              <div style="font-size:1.1rem;font-weight:700;color:var(--brown-dark)">$${total.toFixed(2)}</div>
              ${misprint > 0 ? `<div style="font-size:0.72rem;color:var(--brown-light)">Misprints: $${misprint.toFixed(2)}</div>` : ''}
            </div>
          </div>
        </div>`;
    }).join('');
  } catch (e) {
    container.innerHTML = `<div class="dog-state" style="padding:2rem">${dogEmpty('Could not load sales history')}</div>`;
  }
}

let reportCharts = {};
let reportData = null;

async function renderSalesReportsPage() {
  const today    = new Date().toLocaleDateString('en-CA');
  const monthAgo = new Date(Date.now() - 90 * 86400000).toLocaleDateString('en-CA');
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Sales Reports</h1><p class="page-subtitle">Analyze your performance and revenue</p></div>
      <div class="page-actions">
        <button class="btn btn-secondary" onclick="window.print()">🖨️ Print Report</button>
      </div>
    </div>
    <div class="reports-filters">
      <div style="display:flex;align-items:center;gap:0.5rem;flex-wrap:wrap;">
        <div class="form-field" style="margin:0;min-width:160px;">
          <label class="field-label" style="margin-bottom:0.2rem;">Partner</label>
          <select class="field-input" id="report-partner-filter" style="font-size:0.85rem;">
            <option value="">All Partners</option>
          </select>
        </div>
        <div class="date-range-group">
          <span style="font-size:0.8rem;color:var(--brown-light)">From</span>
          <input type="date" id="date-from" value="${monthAgo}">
          <span style="font-size:0.8rem;color:var(--brown-light)">To</span>
          <input type="date" id="date-to" value="${today}">
        </div>
      </div>
      <div class="period-toggle">
        <button class="period-btn" onclick="setQuickRange(7,this)">7 Days</button>
        <button class="period-btn" onclick="setQuickRange(30,this)">30 Days</button>
        <button class="period-btn active" onclick="setQuickRange(90,this)">3 Months</button>
        <button class="period-btn" onclick="setQuickRange(365,this)">1 Year</button>
        <button class="period-btn" onclick="setQuickRange(0,this)">All Time</button>
      </div>
    </div>
    <div class="stats-grid" style="margin-bottom:1.5rem;">
      ${[
        {icon:'💰',color:'teal', label:'Total Revenue',id:'report-revenue'},
        {icon:'🃏',color:'amber',label:'Cards Sold',     id:'report-sold'},
        {icon:'🏪',color:'green',label:'Top Store',      id:'report-top-store'},
        {icon:'⭐',color:'coral',label:'Top Design',     id:'report-top-design'},
      ].map(s => `
        <div class="stat-card">
          <div class="stat-icon ${s.color}">${s.icon}</div>
          <div class="stat-value" id="${s.id}">—</div>
          <div class="stat-label">${s.label}</div>
        </div>`).join('')}
    </div>
    <div class="stats-grid" style="margin-bottom:1.5rem;">
      ${[
        {icon:'🏪',color:'teal', label:'Retail Revenue',id:'report-retail-rev'},
        {icon:'🎪',color:'amber',label:'Market Revenue',id:'report-market-rev'},
        {icon:'📦',color:'green',label:'Cards on Shelves',id:'report-on-shelves'},
        {icon:'💵',color:'coral',label:'Shelf Value',id:'report-shelf-value'},
      ].map(s => `
        <div class="stat-card">
          <div class="stat-icon ${s.color}">${s.icon}</div>
          <div class="stat-value" id="${s.id}">—</div>
          <div class="stat-label">${s.label}</div>
        </div>`).join('')}
    </div>
    <div class="charts-grid">
      <div class="chart-card chart-card-full"><div class="chart-title">Revenue Over Time</div><div class="chart-container"><canvas id="revenue-chart"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">Sales by Store</div><div class="chart-container"><canvas id="store-chart"></canvas></div></div>
      <div class="chart-card"><div class="chart-title">Top Designs (Est. Units Sold)</div><div class="chart-container"><canvas id="design-chart"></canvas></div></div>
    </div>
    <div id="report-detail-tables" style="margin-top:1.5rem;"></div>`;

  // Load data
  await loadSalesReportData();

  // Set up filter listeners
  document.getElementById('report-partner-filter')?.addEventListener('change', () => loadSalesReportData());
  document.getElementById('date-from')?.addEventListener('change', () => loadSalesReportData());
  document.getElementById('date-to')?.addEventListener('change', () => loadSalesReportData());
}

async function loadSalesReportData() {
  const partnerId = document.getElementById('report-partner-filter')?.value || '';
  const dateFrom = document.getElementById('date-from')?.value || '';
  const dateTo = document.getElementById('date-to')?.value || '';

  try {
    let url = `${GOOGLE_SCRIPT_URL}?action=getSalesReportData`;
    if (partnerId) url += `&partnerId=${partnerId}`;
    if (dateFrom) url += `&dateFrom=${dateFrom}`;
    if (dateTo) url += `&dateTo=${dateTo}`;

    const r = await fetch(url);
    const result = await r.json();
    if (!result.success) throw new Error(result.error);

    reportData = result.data;

    // Populate partner dropdown (only on first load)
    const select = document.getElementById('report-partner-filter');
    if (select && select.options.length <= 1 && reportData.partnerList) {
      reportData.partnerList.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = p.name;
        select.appendChild(opt);
      });
      if (partnerId) select.value = partnerId;
    }

    // Populate stats
    const s = reportData.stats;
    const el = id => document.getElementById(id);
    el('report-revenue').textContent = '$' + s.totalRevenue.toFixed(2);
    el('report-sold').textContent = s.totalCardsSold;
    el('report-top-store').textContent = s.topStore;
    el('report-top-design').textContent = s.topDesign;
    el('report-retail-rev').textContent = '$' + s.totalRetailRevenue.toFixed(2);
    el('report-market-rev').textContent = partnerId ? 'N/A' : '$' + s.totalMarketRevenue.toFixed(2);

    // Shelf stats
    let totalOnShelves = 0, totalShelfValue = 0;
    (reportData.partnerStockSummary || []).forEach(p => {
      totalOnShelves += p.cardsOnShelf;
      totalShelfValue += p.shelfValue;
    });
    el('report-on-shelves').textContent = totalOnShelves;
    el('report-shelf-value').textContent = '$' + totalShelfValue.toFixed(2);

    // Render charts
    renderReportCharts(reportData.charts);

    // Render detail tables
    renderReportDetailTables(reportData);

  } catch (e) {
    showToast('Failed to load report data: ' + e.message, '');
  }
}

function renderReportCharts(charts) {
  if (typeof Chart === 'undefined') return;
  const palette = {
    teal: 'rgba(74,171,171,0.85)',
    amber: 'rgba(232,147,58,0.85)',
    green: 'rgba(90,158,111,0.85)',
    coral: 'rgba(224,92,69,0.85)',
    brown: 'rgba(107,76,59,0.85)',
    colors: ['#4AABAB','#E8933A','#5A9E6F','#E05C45','#6B4C3B','#A07860','#3D2B1F','#C4A882']
  };

  // Destroy old charts
  Object.values(reportCharts).forEach(c => c.destroy());
  reportCharts = {};

  // Revenue Over Time
  const revenueCtx = document.getElementById('revenue-chart')?.getContext('2d');
  if (revenueCtx && charts.monthlyRevenue.length > 0) {
    const labels = charts.monthlyRevenue.map(d => {
      const [y, m] = d.month.split('-');
      return ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(m)] + ' ' + y;
    });
    reportCharts.revenue = new Chart(revenueCtx, {
      type: 'bar',
      data: { labels, datasets: [{ label: 'Revenue', data: charts.monthlyRevenue.map(d => d.revenue), backgroundColor: palette.teal, borderRadius: 6, borderSkipped: false }] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { callback: v => '$' + v } }, x: { grid: { display: false } } } }
    });
  } else if (revenueCtx) {
    reportCharts.revenue = new Chart(revenueCtx, {
      type: 'bar',
      data: { labels: ['No data'], datasets: [{ data: [0], backgroundColor: 'rgba(0,0,0,0.08)' }] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
  }

  // Sales by Store
  const storeCtx = document.getElementById('store-chart')?.getContext('2d');
  if (storeCtx && charts.salesByStore.length > 0) {
    reportCharts.store = new Chart(storeCtx, {
      type: 'doughnut',
      data: {
        labels: charts.salesByStore.map(s => s.name),
        datasets: [{ data: charts.salesByStore.map(s => s.revenue), backgroundColor: palette.colors.slice(0, charts.salesByStore.length), borderWidth: 0 }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } }, cutout: '65%' }
    });
  } else if (storeCtx) {
    reportCharts.store = new Chart(storeCtx, {
      type: 'doughnut',
      data: { labels: ['No data yet'], datasets: [{ data: [1], backgroundColor: ['rgba(0,0,0,0.08)'], borderWidth: 0 }] },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, cutout: '65%' }
    });
  }

  // Top Designs
  const designCtx = document.getElementById('design-chart')?.getContext('2d');
  if (designCtx && charts.topDesigns.length > 0) {
    reportCharts.design = new Chart(designCtx, {
      type: 'bar',
      data: {
        labels: charts.topDesigns.map(d => d.name.length > 20 ? d.name.substring(0,20) + '...' : d.name),
        datasets: [{ label: 'Est. Sold', data: charts.topDesigns.map(d => d.sold), backgroundColor: palette.amber, borderRadius: 6, borderSkipped: false }]
      },
      options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true }, y: { grid: { display: false } } } }
    });
  } else if (designCtx) {
    reportCharts.design = new Chart(designCtx, {
      type: 'bar',
      data: { labels: ['No data'], datasets: [{ data: [0], backgroundColor: 'rgba(0,0,0,0.08)' }] },
      options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } } }
    });
  }
}

function renderReportDetailTables(data) {
  const container = document.getElementById('report-detail-tables');
  if (!container) return;

  // Retail sales history table
  let retailHtml = '';
  if (data.retailSales && data.retailSales.length > 0) {
    const sorted = [...data.retailSales].sort((a, b) => b.month.localeCompare(a.month));
    retailHtml = `
      <div class="card" style="margin-bottom:1.5rem;">
        <h3 class="section-title" style="margin-bottom:0.75rem;">Retail Sales Detail</h3>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
            <thead>
              <tr style="border-bottom:2px solid var(--border);">
                <th style="text-align:left;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Month</th>
                <th style="text-align:left;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Partner</th>
                <th style="text-align:right;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Revenue</th>
                <th style="text-align:right;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Cards</th>
              </tr>
            </thead>
            <tbody>
              ${sorted.map(s => `
                <tr style="border-bottom:1px solid var(--cream-dark);">
                  <td style="padding:0.5rem;">${s.month}</td>
                  <td style="padding:0.5rem;">${s.partnerName}</td>
                  <td style="padding:0.5rem;text-align:right;color:var(--green);font-weight:600;">$${s.revenue.toFixed(2)}</td>
                  <td style="padding:0.5rem;text-align:right;">${s.cardsSold || '—'}</td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  // Market sales history table
  let marketHtml = '';
  if (data.marketSales && data.marketSales.length > 0) {
    const sorted = [...data.marketSales].sort((a, b) => b.date.localeCompare(a.date));
    marketHtml = `
      <div class="card" style="margin-bottom:1.5rem;">
        <h3 class="section-title" style="margin-bottom:0.75rem;">Market & Direct Sales Detail</h3>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
            <thead>
              <tr style="border-bottom:2px solid var(--border);">
                <th style="text-align:left;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Date</th>
                <th style="text-align:left;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Market</th>
                <th style="text-align:right;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Revenue</th>
                <th style="text-align:right;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Cards</th>
              </tr>
            </thead>
            <tbody>
              ${sorted.map(s => `
                <tr style="border-bottom:1px solid var(--cream-dark);">
                  <td style="padding:0.5rem;">${s.date}</td>
                  <td style="padding:0.5rem;">${s.marketName}</td>
                  <td style="padding:0.5rem;text-align:right;color:var(--green);font-weight:600;">$${s.revenue.toFixed(2)}</td>
                  <td style="padding:0.5rem;text-align:right;">${s.cardsSold || '—'}</td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  // Partner stock summary table
  let stockHtml = '';
  if (data.partnerStockSummary && data.partnerStockSummary.length > 0) {
    const sorted = [...data.partnerStockSummary].sort((a, b) => b.shelfValue - a.shelfValue);
    stockHtml = `
      <div class="card" style="margin-bottom:1.5rem;">
        <h3 class="section-title" style="margin-bottom:0.75rem;">Current Stock at Partners</h3>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
            <thead>
              <tr style="border-bottom:2px solid var(--border);">
                <th style="text-align:left;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Partner</th>
                <th style="text-align:right;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Cards on Shelf</th>
                <th style="text-align:right;padding:0.5rem;color:var(--brown-mid);font-size:0.75rem;text-transform:uppercase;">Retail Value</th>
              </tr>
            </thead>
            <tbody>
              ${sorted.map(s => `
                <tr style="border-bottom:1px solid var(--cream-dark);">
                  <td style="padding:0.5rem;">${s.partnerName}</td>
                  <td style="padding:0.5rem;text-align:right;font-weight:600;color:var(--teal);">${s.cardsOnShelf}</td>
                  <td style="padding:0.5rem;text-align:right;color:var(--green);font-weight:600;">$${s.shelfValue.toFixed(2)}</td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  container.innerHTML = retailHtml + marketHtml + stockHtml;
}

function setQuickRange(days, btn) {
  document.querySelectorAll('.period-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  if (days === 0) {
    document.getElementById('date-from').value = '2020-01-01';
  } else {
    document.getElementById('date-from').value = new Date(Date.now() - days * 86400000).toLocaleDateString('en-CA');
  }
  document.getElementById('date-to').value = new Date().toLocaleDateString('en-CA');
  loadSalesReportData();
}

// ============================================================
// INVENTORY AUDITOR
// ============================================================
async function renderInventoryAuditorPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Inventory Audit</h1><p class="page-subtitle">Correct on-hand stock counts after a physical count</p></div>
    </div>
    <div class="card" style="margin-bottom:1.25rem;border-left:4px solid var(--amber);">
      <p style="font-size:0.875rem;color:var(--brown-mid);">⚠️ Use this page after doing a physical count. Enter the actual number you have on hand and save.</p>
    </div>
    <div style="display:flex;align-items:center;gap:0.75rem;margin-bottom:1rem;flex-wrap:wrap;">
      <div class="search-box" style="flex:1;min-width:180px;max-width:360px;">
        <span class="search-icon">⌕</span>
        <input type="text" id="audit-search" placeholder="Search designs...">
      </div>
      <div class="view-toggle" title="Card size">
        <button class="view-toggle-btn ${auditSizeMode === 'comfy' ? 'active' : ''}" id="audit-size-comfy"   onclick="setAuditSize('comfy',this)">▬</button>
        <button class="view-toggle-btn ${auditSizeMode === 'compact' ? 'active' : ''}"         id="audit-size-compact" onclick="setAuditSize('compact',this)">≡</button>
        <button class="view-toggle-btn"         id="audit-size-title"   onclick="setAuditSize('title',this)">☰</button>
      </div>
      <button class="btn btn-primary btn-sm" onclick="saveAudit()" style="margin-left:auto;">✓ Save Updates</button>
    </div>
    <div id="audit-container">${dogLoading('Loading inventory...')}</div>`;

  try {
    const [itemsRes, consignRes] = await Promise.all([
      itemsCache ? Promise.resolve(null) : fetch(`${GOOGLE_SCRIPT_URL}?action=getItems`).then(r => r.json()),
      fetch(`${GOOGLE_SCRIPT_URL}?action=getConsignmentTotals`).then(r => r.json()),
    ]);
    if (itemsRes) {
      if (!itemsRes.success) throw new Error(itemsRes.error);
      itemsCache = itemsRes.items.filter(i => i.ItemID);
    }
    window._consignmentTotals = (consignRes && consignRes.success) ? consignRes.totals : {};
    document.getElementById('audit-search')?.addEventListener('input', (e) => {
      const q        = e.target.value.toLowerCase();
      const filtered = (itemsCache || []).filter(i => (i.DisplayName || i.Name || '').toLowerCase().includes(q) || String(i.ItemID).includes(q));
      renderAuditCards(filtered);
    });
    renderAuditCards(itemsCache);
  } catch (e) {
    document.getElementById('audit-container').innerHTML = dogError(e.message);
  }
}

let auditSizeMode = window.innerWidth <= 1024 ? 'compact' : 'comfy';

function setAuditSize(mode, btn) {
  auditSizeMode = mode;
  document.querySelectorAll('#audit-size-comfy,#audit-size-compact,#audit-size-title').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderAuditCards(itemsCache || []);
}

function renderAuditCard(item) {
  const atStores = (window._consignmentTotals || {})[String(item.ItemID)] || 0;
  const onHand = parseInt(item.StartingAtHome) || 0;
  const total = onHand + atStores;
  return `<div class="audit-card audit-card-${auditSizeMode}">
    <div class="audit-card-info">
      <div class="audit-card-name">${item.DisplayName || item.Name}</div>
      <div class="audit-card-id">#${item.ItemID} · ${item.ProductType || 'Card'}</div>
      <label class="retire-toggle-wrap" title="Mark as Retired">
        <input type="checkbox" class="retire-checkbox" data-item-id="${item.ItemID}" ${item.Status === 'Retired' ? 'checked' : ''}>
        <span class="retire-toggle-label">Retire</span>
      </label>
    </div>
    <div>
      <div class="audit-card-stock" style="color:var(--amber)">${onHand || '—'}</div>
      <div class="audit-card-stock-label">On Hand</div>
    </div>
    <div>
      <div class="audit-card-stock" style="color:var(--teal)">${atStores || '—'}</div>
      <div class="audit-card-stock-label">At Stores</div>
    </div>
    <div>
      <div class="audit-card-stock" style="font-weight:700;color:var(--brown-dark)">${total}</div>
      <div class="audit-card-stock-label">Total</div>
    </div>
    <div class="audit-input-group">
      <input type="number" class="audit-new-input" placeholder="New #" min="0" data-item-id="${item.ItemID}">
      <div class="audit-new-label">Actual Count</div>
    </div>
  </div>`;
}

function renderAuditCards(items) {
  const container = document.getElementById('audit-container');
  if (!container || !items) return;
  const active = items.filter(i => i.Status !== 'Retired');
  const retired = items.filter(i => i.Status === 'Retired');
  container.innerHTML = `
    ${active.map(item => renderAuditCard(item)).join('')}
    <button class="btn btn-primary btn-lg" style="width:100%;margin-top:1rem" onclick="saveAudit()">✓ Save All Updates</button>
    ${retired.length > 0 ? `
      <div style="margin-top:1.5rem;border-top:2px solid var(--brown-light);padding-top:1rem;">
        <div style="display:flex;align-items:center;justify-content:space-between;cursor:pointer;padding:0.5rem 0;" onclick="toggleAuditRetired()">
          <span style="font-weight:600;color:var(--brown-mid);">🪦 Retired Designs (${retired.length})</span>
          <span id="audit-retired-arrow" style="transition:transform 0.2s;">▶</span>
        </div>
        <div id="audit-retired-section" style="display:none;">
          ${retired.map(item => renderAuditCard(item)).join('')}
        </div>
      </div>` : ''}`;
}

function toggleAuditRetired() {
  const section = document.getElementById('audit-retired-section');
  const arrow = document.getElementById('audit-retired-arrow');
  if (!section) return;
  const hidden = section.style.display === 'none';
  section.style.display = hidden ? '' : 'none';
  arrow.textContent = hidden ? '▼' : '▶';
}

async function saveAudit() {
  const inputs     = document.querySelectorAll('.audit-new-input');
  const retireBoxes = document.querySelectorAll('.retire-checkbox');
  const updates    = [];
  const retires    = [];

  inputs.forEach(input => {
    if (input.value !== '') updates.push({ itemId: input.dataset.itemId, newStock: parseInt(input.value) });
  });
  retireBoxes.forEach(cb => {
    const item          = (itemsCache || []).find(i => String(i.ItemID) === String(cb.dataset.itemId));
    const currentStatus = item?.Status || 'Open';
    if (cb.checked && currentStatus !== 'Retired')                    retires.push(cb.dataset.itemId);
    if (!cb.checked && currentStatus === 'Retired') retires.push({ itemId: cb.dataset.itemId, restore: true });
  });

  if (updates.length === 0 && retires.length === 0) { showToast('No changes entered', ''); return; }
  showToast('Saving...', '');

  const stockPromises  = updates.map(u => fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'updateItem', itemData: { itemId: u.itemId, newStock: u.newStock } }) }));
  const retirePromises = retires.map(r => {
    const itemId = typeof r === 'object' ? r.itemId : r;
    const status = typeof r === 'object' && r.restore ? 'Open' : 'Retired';
    return fetch(GOOGLE_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'updateItemStatus', itemId, status }) });
  });

  try {
    await Promise.all([...stockPromises, ...retirePromises]);
    itemsCache = null;
    showToast(`✅ ${updates.length + retires.length} update(s) saved!`, 'success');
    // Re-fetch items so audit cards render with names
    const r = await fetch(`${GOOGLE_SCRIPT_URL}?action=getItems`);
    const d = await r.json();
    if (d.success) {
      itemsCache = d.items.filter(i => i.ItemID);
      renderAuditCards(itemsCache);
    }
  } catch (e) {
    showToast('❌ Some updates failed — try again', '');
  }
}

// ============================================================
// SETTINGS
// ============================================================
function renderSettingsPage() {
  appContainer.innerHTML = `
    <div class="page-header">
      <div><h1 class="page-title">Settings</h1><p class="page-subtitle">App preferences and configuration</p></div>
    </div>
    <div class="settings-section">
      <div class="settings-section-header">🏷️ Business Info</div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Business Name</div><div class="settings-row-desc">Prints by Angel</div></div>
        <button class="btn btn-secondary btn-sm">Edit</button>
      </div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Google Sheet ID</div><div class="settings-row-desc">1FiDZXPV6aimKpKUvzDCQczq01nCdvMZLzRhWq-DB50U</div></div>
        <button class="btn btn-secondary btn-sm" onclick="showToast('Sheet is connected ✅')">Test</button>
      </div>
    </div>
    <div class="settings-section">
      <div class="settings-section-header">🔔 Notifications</div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Low Stock Alerts</div><div class="settings-row-desc">Notify when a design has fewer than 5 cards on hand</div></div>
        <label class="toggle-switch"><input type="checkbox" checked><span class="toggle-slider"></span></label>
      </div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Monthly Summary</div><div class="settings-row-desc">Email yourself a monthly sales summary</div></div>
        <label class="toggle-switch"><input type="checkbox"><span class="toggle-slider"></span></label>
      </div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Vending Machine Removal Reminders</div><div class="settings-row-desc">Notified 7 days, 3 days, 24 hours, and day-of removal</div></div>
        <label class="toggle-switch"><input type="checkbox" checked><span class="toggle-slider"></span></label>
      </div>
    </div>
    <div class="settings-section">
      <div class="settings-section-header">🏷️ Tag Management</div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Manage Tags & Categories</div><div class="settings-row-desc">Add, edit, or organize card tags</div></div>
        <button class="btn btn-secondary btn-sm" onclick="openManageTagsModal()">Manage Tags</button>
      </div>
    </div>
    <div class="settings-section">
      <div class="settings-section-header">📊 Reports</div>
      <div class="settings-row">
        <div class="settings-row-info"><div class="settings-row-label">Google Drive Folder</div><div class="settings-row-desc">Where reports are saved when you tap "Save to Drive"</div></div>
        <button class="btn btn-secondary btn-sm">Set Folder</button>
      </div>
    </div>
    <div style="margin-top:1rem;text-align:center;">
      <p style="font-size:0.75rem;color:var(--brown-light);">Prints by Angel Inventory App · Built with ♥</p>
    </div>`;
}