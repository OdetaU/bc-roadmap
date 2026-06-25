// ── Config ────────────────────────────────────────────────────────────────────
const CONFIG = {
  // Replace with your Azure Functions URL after deployment
  apiBase: window.location.hostname === 'localhost'
    ? 'http://localhost:7071/api'
    : '/api',

  // Azure AD config for login
  msalConfig: {
    auth: {
      clientId: '85682292-7839-4196-a85a-c3f8d196d700',
      authority: 'https://login.microsoftonline.com/70b2dd63-cbac-4385-bed6-cd2cdf4b058f',
      redirectUri: window.location.origin,
    },
    cache: { cacheLocation: 'sessionStorage' },
  },
  scopes: ['User.Read'],
};

// ── Status helpers ─────────────────────────────────────────────────────────────
const STATUS = {
  planned:  { label: 'Planuojama',           cls: 'gray'   },
  notified: { label: 'Klientas informuotas', cls: 'blue'   },
  test:     { label: 'Test aplinkoje',       cls: 'blue'   },
  testing:  { label: 'Testavimas',           cls: 'amber'  },
  prod:     { label: 'Prod aplinkoje',       cls: 'purple' },
  done:     { label: 'Baigta',               cls: 'green'  },
};

const TYPE = { appreg: 'App Reg. Secret', admin: 'Admin Relationship' };

function daysUntil(d) { return d ? Math.ceil((new Date(d) - new Date()) / 86400000) : Infinity; }
function fmt(d) { if (!d) return '—'; const [y,m,dd] = d.slice(0,10).split('-'); return `${dd}.${m}.${y}`; }
function clientName(id) { return (App.state.clients.find(c => c.id === id || c.tenantId === id) || {}).name || id || '—'; }

// ── UI helpers ────────────────────────────────────────────────────────────────
const UI = {
  closeModal(id) { document.getElementById(id).classList.remove('open'); },
  openModal(id)  { document.getElementById(id).classList.add('open'); },
  toast(msg, dur = 2500) {
    const t = document.getElementById('toast');
    t.textContent = msg; t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), dur);
  },
  loading(sectionId) {
    document.getElementById(sectionId).innerHTML = '<div class="loading">Kraunama...</div>';
  },
};

// ── API calls ─────────────────────────────────────────────────────────────────
async function api(path, method = 'GET', body) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (body) opts.body = JSON.stringify(body);
  const r = await fetch(CONFIG.apiBase + path, opts);
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}

// ── App ───────────────────────────────────────────────────────────────────────
const App = {
  state: { clients: [], updates: [], secrets: [], versions: [], user: null },
  editUpdateSpId: null,
  editSecretSpId: null,
  editUpdateClientId: null,
  editSecretClientId: null,

  // ── Auth (simple — MSAL loaded from CDN) ──────────────────────────────────
  msalInstance: null,

  async initAuth() {
    // Load MSAL dynamically
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.18.0/lib/msal-browser.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
    this.msalInstance = new msal.PublicClientApplication(CONFIG.msalConfig);
    await this.msalInstance.initialize();

    // Handle redirect
    const resp = await this.msalInstance.handleRedirectPromise();
    if (resp) {
      this.state.user = resp.account;
      this.showApp();
      return;
    }
    const accounts = this.msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      this.state.user = accounts[0];
      this.showApp();
    }
    // Otherwise stay on login screen
  },

  login() {
    this.msalInstance.loginRedirect({ scopes: CONFIG.scopes });
  },

  logout() {
    this.msalInstance.logoutRedirect();
  },

  showApp() {
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app').style.display = 'block';
    const u = this.state.user;
    document.getElementById('user-name').textContent = u.name || u.username || '';
    this.refresh();
  },

  async refresh() {
    try {
      const [clients, updates, secrets, versions] = await Promise.all([
        api('/getClients'),
        api('/getUpdates'),
        api('/getSecrets'),
        api('/getVersions'),
      ]);
      this.state.clients  = clients;
      this.state.updates  = updates;
      this.state.secrets  = secrets;
      this.state.versions = versions;
      this.renderCurrentTab();
      UI.toast('Duomenys atnaujinti');
    } catch (e) {
      UI.toast('Klaida: ' + e.message, 4000);
    }
  },

  currentTab: 'dashboard',
  showTab(t) {
    this.currentTab = t;
    document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-' + t).classList.add('active');
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.getElementById('sec-' + t).classList.add('active');
    this.renderCurrentTab();
  },
  renderCurrentTab() {
    const t = this.currentTab;
    if (t === 'dashboard') this.renderDashboard();
    else if (t === 'updates') this.renderUpdates();
    else if (t === 'gantt') this.renderGantt();
    else if (t === 'secrets') this.renderSecrets();
    else if (t === 'versions') this.renderVersions();
  },

  // ── Dashboard ──────────────────────────────────────────────────────────────
  renderDashboard() {
    const sec = document.getElementById('sec-dashboard');
    const { clients, updates, secrets } = this.state;
    const active  = updates.filter(u => !['done','planned'].includes(u.status));
    const planned = updates.filter(u => u.status === 'planned');
    const done    = updates.filter(u => u.status === 'done');
    const expiredS  = secrets.filter(s => daysUntil(s.expires) < 0);
    const warningS  = secrets.filter(s => { const d = daysUntil(s.expires); return d >= 0 && d <= (s.warn || 30); });
    const notifyDue = updates.filter(u => { const d = daysUntil(u.dNotify); return d >= 0 && d <= 7 && u.status === 'planned'; });

    let alerts = '';
    expiredS.forEach(s => {
      alerts += `<div class="alert-row"><strong class="alert-label">Pasibaigę!</strong> <span><b>${clientName(s.clientId)}</b> — ${s.name} baigėsi ${fmt(s.expires)}</span></div>`;
    });
    warningS.forEach(s => {
      alerts += `<div class="alert-row warn"><strong class="alert-label warn">Artėja pabaiga</strong> <span><b>${clientName(s.clientId)}</b> — ${s.name} baigiasi po <b>${Math.round(daysUntil(s.expires))}</b> d. (${fmt(s.expires)})</span></div>`;
    });
    notifyDue.forEach(u => {
      alerts += `<div class="alert-row warn"><strong class="alert-label warn">Reikia informuoti!</strong> <span><b>${clientName(u.clientId)}</b> — naujinimas ${u.newVer}, informuoti iki ${fmt(u.dNotify)}</span></div>`;
    });

    const upcoming = [...updates].filter(u => u.status !== 'done').sort((a,b) => (a.dTest||'9') > (b.dTest||'9') ? 1 : -1);

    sec.innerHTML = `
    <div class="summary-cards">
      <div class="summary-card"><div class="num">${clients.length}</div><div class="lbl">Klientai</div></div>
      <div class="summary-card"><div class="num">${active.length}</div><div class="lbl">Aktyvūs naujinimai</div></div>
      <div class="summary-card"><div class="num">${planned.length}</div><div class="lbl">Planuojama</div></div>
      <div class="summary-card"><div class="num">${done.length}</div><div class="lbl">Baigta</div></div>
      <div class="summary-card"><div class="num" style="${expiredS.length ? 'color:var(--red)' : ''}">${expiredS.length + warningS.length}</div><div class="lbl">Secrets įspėjimai</div></div>
    </div>
    ${alerts || '<div style="color:var(--text2);font-size:13px;margin-bottom:1.5rem">✓ Nėra aktyvių įspėjimų</div>'}
    <div class="section-title">Artimiausi naujinimai</div>
    <div class="card"><table>
      <thead><tr><th>Klientas</th><th>Naujinimas</th><th>Test data</th><th>Prod data</th><th>Statusas</th></tr></thead>
      <tbody>${upcoming.map(u => {
        const si = STATUS[u.status] || STATUS.planned;
        return `<tr>
          <td class="client-name">${clientName(u.clientId)}</td>
          <td>${u.curVer} → <b>${u.newVer}</b></td>
          <td>${fmt(u.dTest)}</td>
          <td>${fmt(u.dProd)}</td>
          <td><span class="badge ${si.cls}">${si.label}</span></td>
        </tr>`;
      }).join('') || '<tr><td colspan="5" class="empty">Nėra naujinimų</td></tr>'}</tbody>
    </table></div>`;
  },

  // ── Updates ────────────────────────────────────────────────────────────────
  renderUpdates() {
    const sec = document.getElementById('sec-updates');
    const filterVal = document.getElementById('filter-update-status')?.value || '';
    const rows = this.state.updates
      .filter(u => !filterVal || u.status === filterVal)
      .sort((a,b) => (a.dTest||'9') > (b.dTest||'9') ? 1 : -1);

    sec.innerHTML = `
    <div class="toolbar">
      <button class="btn primary" onclick="App.openUpdateModal()">+ Pridėti naujinimą</button>
      <select id="filter-update-status" onchange="App.renderUpdates()" style="font-size:13px;padding:6px 10px;border-radius:var(--radius);border:0.5px solid var(--border2);background:var(--bg);color:var(--text)">
        <option value="">Visi statusai</option>
        ${Object.entries(STATUS).map(([v,i]) => `<option value="${v}" ${filterVal===v?'selected':''}>${i.label}</option>`).join('')}
      </select>
    </div>
    <div class="card"><table>
      <thead><tr><th>Klientas</th><th>Versija</th><th>Aplinka</th><th>Kėlimas į Test</th><th>Testavimo pabaiga</th><th>Kėlimas į Prod</th><th>Informuoti iki</th><th>Statusas</th><th></th></tr></thead>
      <tbody>${rows.map(u => {
        const si = STATUS[u.status] || STATUS.planned;
        const nd = daysUntil(u.dNotify);
        const nw = nd >= 0 && nd <= 7 && u.status === 'planned';
        return `<tr>
          <td class="client-name">${clientName(u.clientId)}</td>
          <td>${u.curVer} → <b>${u.newVer}</b></td>
          <td><span class="badge gray">${u.env}</span></td>
          <td>${fmt(u.dTest)}</td>
          <td>${fmt(u.dTestEnd)}</td>
          <td>${fmt(u.dProd)}</td>
          <td style="${nw ? 'color:var(--amber);font-weight:500' : ''}">${fmt(u.dNotify)}${nw ? ' ⚠' : ''}</td>
          <td><span class="badge ${si.cls}">${si.label}</span></td>
          <td style="white-space:nowrap">
            <button class="icon-btn" onclick="App.openUpdateModal('${u.spItemId||''}','${u.clientId}','${u.curVer}','${u.newVer}','${u.env}','${u.dTest||''}','${u.dTestEnd||''}','${u.dProd||''}','${u.dNotify||''}','${u.status}',\`${(u.notes||'').replace(/`/g,"'")}\`)">✎</button>
            <button class="icon-btn danger" onclick="App.deleteUpdate('${u.spItemId}')">✕</button>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="9" class="empty">Nėra įrašų</td></tr>'}</tbody>
    </table></div>`;
  },

  openUpdateModal(spId, clientId, curVer, newVer, env, dTest, dTestEnd, dProd, dNotify, status, notes) {
    this.editUpdateSpId = spId || null;
    this.editUpdateClientId = clientId || null;
    document.getElementById('modal-update-title').textContent = spId ? 'Redaguoti naujinimą' : 'Naujas naujinimas';
    const sel = document.getElementById('f-client');
    sel.innerHTML = this.state.clients.map(c => `<option value="${c.id || c.tenantId}">${c.name}</option>`).join('');
    if (clientId) sel.value = clientId;
    document.getElementById('f-cur-ver').value = curVer || '';
    document.getElementById('f-new-ver').value = newVer || '';
    document.getElementById('f-env').value = env || 'Sandbox + Prod';
    document.getElementById('f-d-test').value = dTest || '';
    document.getElementById('f-d-testend').value = dTestEnd || '';
    document.getElementById('f-d-prod').value = dProd || '';
    document.getElementById('f-d-notify').value = dNotify || '';
    document.getElementById('f-status').value = status || 'planned';
    document.getElementById('f-notes').value = notes || '';
    UI.openModal('modal-update');
  },

  async saveUpdate() {
    const body = {
      spItemId:  this.editUpdateSpId,
      clientId:  document.getElementById('f-client').value,
      curVer:    document.getElementById('f-cur-ver').value,
      newVer:    document.getElementById('f-new-ver').value,
      env:       document.getElementById('f-env').value,
      dTest:     document.getElementById('f-d-test').value,
      dTestEnd:  document.getElementById('f-d-testend').value,
      dProd:     document.getElementById('f-d-prod').value,
      dNotify:   document.getElementById('f-d-notify').value,
      status:    document.getElementById('f-status').value,
      notes:     document.getElementById('f-notes').value,
    };
    try {
      await api('/saveUpdate', 'POST', body);
      UI.closeModal('modal-update');
      UI.toast('Išsaugota');
      await this.refresh();
    } catch (e) { UI.toast('Klaida: ' + e.message, 4000); }
  },

  async deleteUpdate(spId) {
    if (!spId || !confirm('Ištrinti šį naujinimą?')) return;
    try {
      await api('/deleteUpdate/' + spId, 'DELETE');
      UI.toast('Ištrinta');
      await this.refresh();
    } catch (e) { UI.toast('Klaida: ' + e.message, 4000); }
  },

  // ── Gantt ──────────────────────────────────────────────────────────────────
  renderGantt() {
    const sec = document.getElementById('sec-gantt');
    const updates = this.state.updates;
    const allDates = updates.flatMap(u => [u.dTest, u.dTestEnd, u.dProd].filter(Boolean));
    if (!allDates.length) { sec.innerHTML = '<div class="empty">Nėra naujinimų su nustatytomis datomis</div>'; return; }

    const minD = new Date(allDates.reduce((a,b) => a < b ? a : b));
    const maxD = new Date(allDates.reduce((a,b) => a > b ? a : b));
    minD.setDate(minD.getDate() - 3);
    maxD.setDate(maxD.getDate() + 3);

    const days = [];
    const cur = new Date(minD);
    while (cur <= maxD) { days.push(new Date(cur)); cur.setDate(cur.getDate() + 1); }

    const months = [];
    days.forEach(d => {
      const key = `${d.getFullYear()}-${d.getMonth()}`;
      if (!months.length || months[months.length-1].key !== key)
        months.push({ key, label: d.toLocaleString('lt', { month: 'long', year: 'numeric' }), count: 0 });
      months[months.length-1].count++;
    });

    const todayIdx = days.findIndex(d => d.toDateString() === new Date().toDateString());

    let html = `<div class="gantt-wrap"><table class="gantt-table">
    <thead>
      <tr>
        <th class="gantt-label">Klientas / Naujinimas</th>
        ${months.map(m => `<th colspan="${m.count}" style="text-align:center">${m.label}</th>`).join('')}
      </tr>
      <tr>
        <th class="gantt-label"></th>
        ${days.map((d,i) => `<th class="gantt-cell ${i===todayIdx?'today-header':''}" style="text-align:center;font-size:10px;padding:4px 2px">${d.getDate()}</th>`).join('')}
      </tr>
    </thead><tbody>`;

    updates.sort((a,b) => (a.dTest||'9') > (b.dTest||'9') ? 1 : -1).forEach(u => {
      const si = STATUS[u.status] || STATUS.planned;
      html += `<tr><td class="gantt-label"><b>${clientName(u.clientId)}</b>
        <div class="sub">${u.curVer} → ${u.newVer}</div>
        <div style="margin-top:3px"><span class="badge ${si.cls}" style="font-size:10px">${si.label}</span></div>
      </td>`;

      days.forEach((d, i) => {
        const ds = d.toISOString().slice(0, 10);
        let bar = '';
        if (u.dTest && ds === u.dTest) {
          const span = u.dTestEnd ? Math.ceil((new Date(u.dTestEnd) - new Date(u.dTest)) / 86400000) + 1 : 1;
          bar = `<div class="gantt-bar test-period" style="left:1px;${span>1?`width:${span*26-2}px`:'right:1px'}" title="${clientName(u.clientId)}: Test + testavimas">${span>4?'Test':''}</div>`;
        }
        if (u.dProd && ds === u.dProd) {
          bar = `<div class="gantt-bar prod-day" style="left:1px;right:1px" title="${clientName(u.clientId)}: Kėlimas į Prod"></div>`;
        }
        html += `<td class="gantt-cell ${i===todayIdx?'today-col':''}">${bar}</td>`;
      });
      html += '</tr>';
    });

    html += `</tbody></table></div>
    <div class="legend" style="margin-top:10px">
      <span class="legend-item"><span class="legend-dot" style="background:#185FA5"></span>Kėlimas į Test + testavimo laikotarpis</span>
      <span class="legend-item"><span class="legend-dot" style="background:#3B6D11"></span>Kėlimas į Prod</span>
      <span class="legend-item"><span class="legend-dot" style="background:rgba(24,95,165,0.15)"></span>Šiandiena</span>
    </div>`;

    sec.innerHTML = html;
  },

  // ── Secrets ────────────────────────────────────────────────────────────────
  renderSecrets() {
    const sec = document.getElementById('sec-secrets');
    const fType   = document.getElementById('filter-stype')?.value || '';
    const fClient = document.getElementById('filter-sclient')?.value || '';
    const rows = this.state.secrets
      .filter(s => (!fType || s.type === fType) && (!fClient || s.clientId === fClient))
      .sort((a,b) => (a.expires||'9') > (b.expires||'9') ? 1 : -1);

    sec.innerHTML = `
    <div class="toolbar">
      <button class="btn primary" onclick="App.openSecretModal()">+ Pridėti rankinį įrašą</button>
      <select id="filter-stype" onchange="App.renderSecrets()" style="font-size:13px;padding:6px 10px;border-radius:var(--radius);border:0.5px solid var(--border2);background:var(--bg);color:var(--text)">
        <option value="">Visi tipai</option>
        <option value="appreg" ${fType==='appreg'?'selected':''}>App Reg. Secret</option>
        <option value="admin" ${fType==='admin'?'selected':''}>Admin Relationship</option>
      </select>
      <select id="filter-sclient" onchange="App.renderSecrets()" style="font-size:13px;padding:6px 10px;border-radius:var(--radius);border:0.5px solid var(--border2);background:var(--bg);color:var(--text)">
        <option value="">Visi klientai</option>
        ${this.state.clients.map(c => `<option value="${c.id||c.tenantId}" ${fClient===(c.id||c.tenantId)?'selected':''}>${c.name}</option>`).join('')}
      </select>
    </div>
    <div class="card"><table>
      <thead><tr><th>Klientas</th><th>Tipas</th><th>Pavadinimas</th><th>Šaltinis</th><th>Galioja iki</th><th>Liko dienų</th><th>Statusas</th><th></th></tr></thead>
      <tbody>${rows.map(s => {
        const d = daysUntil(s.expires);
        let badge, dText;
        if (d < 0)         { badge = '<span class="badge red">Baigėsi</span>';  dText = `<span style="color:var(--red)">${Math.abs(Math.round(d))} d. atgal</span>`; }
        else if (d <= (s.warn||30)) { badge = '<span class="badge amber">Artėja</span>';  dText = `<span style="color:var(--amber)">${Math.round(d)}</span>`; }
        else               { badge = '<span class="badge green">Galioja</span>'; dText = Math.round(d); }
        const srcBadge = s.source === 'manual' ? '<span class="badge gray">Rankinis</span>' : s.source === 'appreg' ? '<span class="badge blue">Auto</span>' : '<span class="badge purple">Auto</span>';
        return `<tr>
          <td class="client-name">${clientName(s.clientId)}</td>
          <td><span class="badge ${s.type==='appreg'?'blue':'purple'}">${TYPE[s.type]||s.type}</span></td>
          <td>${s.name}${s.notes?`<div class="sub">${s.notes}</div>`:''}</td>
          <td>${srcBadge}</td>
          <td>${fmt(s.expires)}</td>
          <td>${dText}</td>
          <td>${badge}</td>
          <td style="white-space:nowrap">${s.source==='manual'?`
            <button class="icon-btn" onclick="App.openSecretModal('${s.spItemId||''}','${s.clientId}','${s.type}',\`${s.name}\`,'${s.expires}','${s.warn||30}',\`${(s.notes||'').replace(/`/g,"'")}\`)">✎</button>
            <button class="icon-btn danger" onclick="App.deleteSecret('${s.spItemId}')">✕</button>
          `:'<button class="icon-btn" onclick="App.openSecretModal(\''+s.spItemId+'\',\''+s.clientId+'\',\''+s.type+'\',`'+s.name+'`,\''+s.expires+'\',\'${s.warn||30}\',\'\')">✎</button>'}</td>
        </tr>`;
      }).join('') || '<tr><td colspan="8" class="empty">Nėra įrašų</td></tr>'}</tbody>
    </table></div>`;
  },

  openSecretModal(spId, clientId, type, name, expires, warn, notes) {
    this.editSecretSpId = spId || null;
    document.getElementById('modal-secret-title').textContent = spId ? 'Redaguoti įrašą' : 'Naujas rankinis įrašas';
    const sel = document.getElementById('fs-client');
    sel.innerHTML = this.state.clients.map(c => `<option value="${c.id||c.tenantId}">${c.name}</option>`).join('');
    if (clientId) sel.value = clientId;
    document.getElementById('fs-type').value = type || 'appreg';
    document.getElementById('fs-name').value = name || '';
    document.getElementById('fs-expires').value = expires || '';
    document.getElementById('fs-warn').value = warn || 30;
    document.getElementById('fs-notes').value = notes || '';
    UI.openModal('modal-secret');
  },

  async saveSecret() {
    const body = {
      spItemId: this.editSecretSpId,
      clientId: document.getElementById('fs-client').value,
      type:     document.getElementById('fs-type').value,
      name:     document.getElementById('fs-name').value,
      expires:  document.getElementById('fs-expires').value,
      warn:     parseInt(document.getElementById('fs-warn').value) || 30,
      notes:    document.getElementById('fs-notes').value,
      source:   'manual',
    };
    try {
      await api('/saveSecret', 'POST', body);
      UI.closeModal('modal-secret');
      UI.toast('Išsaugota');
      await this.refresh();
    } catch (e) { UI.toast('Klaida: ' + e.message, 4000); }
  },

  async deleteSecret(spId) {
    if (!spId || !confirm('Ištrinti šį įrašą?')) return;
    try {
      await api('/deleteSecret/' + spId, 'DELETE');
      UI.toast('Ištrinta');
      await this.refresh();
    } catch (e) { UI.toast('Klaida: ' + e.message, 4000); }
  },

  // ── BC Versions ────────────────────────────────────────────────────────────
  renderVersions() {
    const sec = document.getElementById('sec-versions');
    const byTenant = {};
    this.state.versions.forEach(v => {
      if (!byTenant[v.tenantId]) byTenant[v.tenantId] = { name: v.clientName, envs: [] };
      byTenant[v.tenantId].envs.push(v);
    });

    sec.innerHTML = `
    <div class="section-title">BC aplinkų versijos (tiesiai iš BC Admin API)</div>
    <div class="versions-grid">
      ${Object.entries(byTenant).map(([tid, t]) => `
        <div class="version-card">
          <div class="version-card-header">${t.name || tid}</div>
          <div class="sub" style="margin-bottom:8px">${tid}</div>
          ${t.envs.map(e => `
            <div class="env-row">
              <span>${e.envName} <span class="badge ${e.envType==='Production'?'green':'blue'}" style="font-size:10px">${e.envType}</span></span>
              <span style="font-family:monospace;font-size:12px">${e.version || '—'}</span>
            </div>
          `).join('')}
        </div>
      `).join('') || '<div class="empty">Nėra duomenų iš BC Admin API</div>'}
    </div>`;
  },
};

// ── Close modals on backdrop click ────────────────────────────────────────────
document.querySelectorAll('.modal-bg').forEach(m => {
  m.addEventListener('click', e => { if (e.target === m) m.classList.remove('open'); });
});

// ── Init ──────────────────────────────────────────────────────────────────────
App.initAuth();
