/* Planificare Vizite - Full rewrite with:
   - Local auth: user+pass, roles: agent/coordonator
   - Clients: county, phone, address, contacts (dynamic)
   - Client "page" as modal (edit full details + contacts + report placeholders)
   - Excel import w/ dedupe (Name+County fallback Name)
   - Visits planning, role filtering (agent sees only own)
   - Export Excel format: Zi | Data | Client | Activitate | Detalii | Observatii
   - Persist localStorage
*/

const LS_KEY = "pv_state_v2";
const SESSION_KEY = "pv_session_v2";

const dayNames = ["DUMINICA","LUNI","MARTI","MIERCURI","JOI","VINERI","SAMBATA"]; // fără diacritice

let state = loadState();
let session = loadSession();

/** Session:
 * { userId, username, role, agentId? }
 */

document.addEventListener("DOMContentLoaded", () => {
  bindTabs();

  // Auth
  $("#btnLogin").addEventListener("click", onLogin);
  $("#btnLogout").addEventListener("click", onLogout);

  // Global
  $("#btnReset").addEventListener("click", onResetAll);
  $("#btnExport").addEventListener("click", onExport);

  // Clients
  $("#btnImportClients").addEventListener("click", onImportClients);
  $("#btnAddClient").addEventListener("click", onAddClient);
  $("#btnClearClients").addEventListener("click", onClearClients);
  $("#clientSearch").addEventListener("input", renderClients);

  // Agents & Routes
  $("#btnAddAgent").addEventListener("click", onAddAgent);
  $("#btnAddRoute").addEventListener("click", onAddRoute);

  // Activities
  $("#btnAddActivity").addEventListener("click", onAddActivity);

  // Planning
  $("#btnGenerate").addEventListener("click", onGenerateVisits);
  $("#btnClearVisits").addEventListener("click", onClearVisitsInInterval);
  $("#btnAddVisit").addEventListener("click", onAddVisitManual);
  $("#btnSort").addEventListener("click", onSortVisits);
  $("#agentFilter").addEventListener("change", () => renderAll());

  // Reports
  $("#btnRefreshReport").addEventListener("click", renderReport);

  // Client modal
  $("#btnCloseClientModal").addEventListener("click", closeClientModal);
  $("#btnSaveClientModal").addEventListener("click", saveClientModal);
  $("#btnDeleteClientModal").addEventListener("click", deleteClientFromModal);
  $("#btnAddContact").addEventListener("click", addContactInModal);
  $("#cmMinus").addEventListener("click", () => changeClientModalCount(-1));
  $("#cmPlus").addEventListener("click", () => changeClientModalCount(+1));

  // Users (coord only)
  $("#btnCreateUser").addEventListener("click", onCreateUser);

  // Defaults
  if (!state.activities.length) {
    state.activities = [
      { id: uid(), name: "comanda" },
      { id: uid(), name: "incasare" },
      { id: uid(), name: "comanda + incasare" },
    ];
  }

  // Default dates: current month
  const now = new Date();
  const from = new Date(now.getFullYear(), now.getMonth(), 1);
  const to = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  $("#fromDate").value = toISO(from);
  $("#toDate").value = toISO(to);
  $("#visitDate").value = toISO(now);

  // Ensure at least one coordinator user exists
  ensureDefaultAdmin();

  saveState(state);
  updateAuthUI();
  renderAll();
});

/* ===================== AUTH ===================== */

function ensureDefaultAdmin() {
  if (!state.users) state.users = [];
  if (state.users.some(u => u.username === "admin")) return;

  state.users.push({
    id: uid(),
    username: "admin",
    passHash: hash("admin123"),
    role: "coordonator",
    agentId: null
  });
}

function onLogin() {
  const username = $("#loginUser").value.trim();
  const pass = $("#loginPass").value;
  const err = $("#loginError");
  err.textContent = "";

  if (!username || !pass) {
    err.textContent = "Completează utilizator și parolă.";
    return;
  }

  const user = (state.users || []).find(u => u.username === username);
  if (!user || user.passHash !== hash(pass)) {
    err.textContent = "Utilizator sau parolă greșită.";
    return;
  }

  session = {
    userId: user.id,
    username: user.username,
    role: user.role,
    agentId: user.agentId || null
  };
  saveSession(session);
  updateAuthUI();
  renderAll();
}

function onLogout() {
  session = null;
  saveSession(null);
  updateAuthUI();
}

function updateAuthUI() {
  const overlay = $("#authOverlay");
  const who = $("#whoami");

  if (!session) {
    overlay.classList.remove("hidden");
    who.textContent = "Neautentificat";
    disableApp(true);
    return;
  }

  overlay.classList.add("hidden");
  disableApp(false);

  const role = session.role;
  if (role === "agent") {
    const agentName = getAgentName(session.agentId);
    who.textContent = `User: ${session.username} • Rol: agent • Agent: ${agentName || "—"}`;
  } else {
    who.textContent = `User: ${session.username} • Rol: coordonator`;
  }

  // Tab Users only for coordinator
  $("#tabUsersBtn").style.display = (session.role === "coordonator") ? "" : "none";
}

function disableApp(disabled) {
  document.querySelectorAll(".topbar button, .tabs button, main input, main select, main button")
    .forEach(el => {
      if (el.id === "btnLogin" || el.id === "loginUser" || el.id === "loginPass") return;
      el.disabled = disabled;
    });
}

/* ===================== RENDER ===================== */

function renderAll() {
  if (!session) return;

  // role gates
  const isCoord = session.role === "coordonator";
  // hide agents management for agent (still can view list, but prevent edits)
  // We'll handle by disabling relevant buttons/inputs in those sections.
  gateCoordOnlyUI();

  renderClients();
  renderAgents();
  renderRoutes();
  renderActivities();
  renderPlanningSelectors();
  renderAgentFilter();
  renderVisits();
  renderReport();
  renderUsers();
}

function gateCoordOnlyUI() {
  const isCoord = session.role === "coordonator";

  // Agents tab create buttons
  $("#btnAddAgent").style.display = isCoord ? "" : "none";
  $("#btnAddRoute").style.display = isCoord ? "" : "none";
  $("#newAgentName").disabled = !isCoord;
  $("#newRouteName").disabled = !isCoord;
  $("#newRouteAgent").disabled = !isCoord;

  // Import/clear clients permitted for coordinator; agent can view+edit only clients used? (we keep simple: agent can view clients but not mass import/delete)
  $("#btnImportClients").style.display = isCoord ? "" : "none";
  $("#fileClients").disabled = !isCoord;
  $("#sheetName").disabled = !isCoord;
  $("#btnClearClients").style.display = isCoord ? "" : "none";

  // Users tab (already hidden for agent)
  // Reset full data only coordinator
  $("#btnReset").style.display = isCoord ? "" : "none";
}

/* ===================== TABS ===================== */
function bindTabs() {
  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      $("#" + btn.dataset.tab).classList.add("active");
    });
  });
}

/* ===================== CLIENTS ===================== */

function onAddClient() {
  if (session.role !== "coordonator") return alert("Doar coordonatorul poate adăuga clienți manual în MVP.");

  const name = $("#newClientName").value.trim();
  const county = $("#newClientCounty").value.trim();
  const address = $("#newClientAddress").value.trim();
  const phone = $("#newClientPhone").value.trim();
  if (!name) return alert("Completează numele clientului.");

  state.clients.push(newClient({ name, county, address, phone }));
  $("#newClientName").value = "";
  $("#newClientCounty").value = "";
  $("#newClientAddress").value = "";
  $("#newClientPhone").value = "";
  $("#clientSearch").value = "";
  saveRender();
}

function onClearClients() {
  if (session.role !== "coordonator") return;
  if (!confirm("Ștergi toți clienții?")) return;
  state.clients = [];
  // also remove visits
  state.visits = [];
  saveRender();
}

function renderClients() {
  const q = ($("#clientSearch").value || "").trim().toLowerCase();
  const tbody = $("#tblClients tbody");
  tbody.innerHTML = "";

  const list = (state.clients || [])
    .filter(c => {
      if (!q) return true;
      return (c.name || "").toLowerCase().includes(q)
        || (c.county || "").toLowerCase().includes(q)
        || (c.phone || "").toLowerCase().includes(q)
        || (c.address || "").toLowerCase().includes(q);
    })
    .sort((a,b)=> (a.name||"").localeCompare(b.name||""));

  for (const c of list) {
    const tr = document.createElement("tr");

    // Name (click opens modal)
    const tdName = document.createElement("td");
    const a = document.createElement("button");
    a.className = "btn ghost";
    a.style.padding = "8px 10px";
    a.textContent = c.name || "";
    a.addEventListener("click", () => openClientModal(c.id));
    tdName.appendChild(a);
    tr.appendChild(tdName);

    // County (editable inline for coordinator)
    tr.appendChild(tdInlineEdit(c.county || "", v => { c.county = v; saveRender(); }, session.role === "coordonator"));

    // Phone
    tr.appendChild(tdInlineEdit(c.phone || "", v => { c.phone = v; saveRender(); }, session.role === "coordonator"));

    // Address
    tr.appendChild(tdInlineEdit(c.address || "", v => { c.address = v; saveRender(); }, session.role === "coordonator"));

    // Visits/month stepper
    const tdFreq = document.createElement("td");
    tdFreq.appendChild(buildStepper(c.monthlyCount ?? 0, (next) => {
      if (session.role !== "coordonator") return;
      c.monthlyCount = clampInt(next, 0, 50);
      saveRender();
    }, session.role === "coordonator"));
    tr.appendChild(tdFreq);

    // Weeks
    const tdWeeks = document.createElement("td");
    tdWeeks.appendChild(weeksPicker(c.monthlyWeeks || [], weeks => {
      if (session.role !== "coordonator") return;
      c.monthlyWeeks = weeks;
      saveRender();
    }, session.role === "coordonator"));
    tr.appendChild(tdWeeks);

    // Delete
    const tdDel = document.createElement("td");
    if (session.role === "coordonator") {
      const btn = document.createElement("button");
      btn.className = "btn danger";
      btn.textContent = "Șterge";
      btn.addEventListener("click", () => {
        if (!confirm(`Ștergi clientul "${c.name}"?`)) return;
        state.clients = state.clients.filter(x => x.id !== c.id);
        state.visits = state.visits.filter(v => v.clientId !== c.id);
        saveRender();
      });
      tdDel.appendChild(btn);
    } else {
      tdDel.textContent = "—";
      tdDel.className = "muted";
    }
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

function newClient({ name, county="", address="", phone="" }) {
  return {
    id: uid(),
    name,
    county,
    address,
    phone,
    freqType: "monthly",
    monthlyCount: 2,
    monthlyWeeks: [1,3],
    contacts: [],
    reports: {} // placeholder for future integrations
  };
}

function buildStepper(value, onChange, enabled=true) {
  const wrap = document.createElement("div");
  wrap.className = "stepper";

  const minus = document.createElement("button");
  minus.type = "button";
  minus.className = "stepBtn";
  minus.textContent = "−";
  minus.disabled = !enabled;

  const val = document.createElement("div");
  val.className = "stepVal";
  val.textContent = String(value);

  const plus = document.createElement("button");
  plus.type = "button";
  plus.className = "stepBtn";
  plus.textContent = "+";
  plus.disabled = !enabled;

  const set = (n) => { val.textContent = String(n); onChange(n); };

  minus.addEventListener("click", () => set((parseInt(val.textContent,10)||0) - 1));
  plus.addEventListener("click", () => set((parseInt(val.textContent,10)||0) + 1));

  // hold for faster
  let holdTimer = null;
  function startHold(delta){
    if (!enabled) return;
    stopHold();
    holdTimer = setInterval(() => set((parseInt(val.textContent,10)||0) + delta), 140);
  }
  function stopHold(){ if (holdTimer) clearInterval(holdTimer); holdTimer = null; }

  minus.addEventListener("pointerdown", () => startHold(-1));
  plus.addEventListener("pointerdown", () => startHold(+1));
  minus.addEventListener("pointerup", stopHold);
  plus.addEventListener("pointerup", stopHold);
  minus.addEventListener("pointerleave", stopHold);
  plus.addEventListener("pointerleave", stopHold);

  wrap.appendChild(minus);
  wrap.appendChild(val);
  wrap.appendChild(plus);
  return wrap;
}

function weeksPicker(selected, onChange, enabled=true) {
  const wrap = document.createElement("div");
  wrap.className = "row wrap";
  for (let w = 1; w <= 5; w++) {
    const chip = document.createElement("div");
    chip.className = "chip" + (selected.includes(w) ? " on" : "");
    chip.textContent = `S${w}`;
    chip.style.opacity = enabled ? "1" : "0.65";
    chip.style.pointerEvents = enabled ? "auto" : "none";
    chip.addEventListener("click", () => {
      const set = new Set(selected);
      if (set.has(w)) set.delete(w); else set.add(w);
      onChange([...set].sort((a,b)=>a-b));
    });
    wrap.appendChild(chip);
  }
  return wrap;
}

function tdInlineEdit(value, onChange, enabled=true) {
  const td = document.createElement("td");
  if (!enabled) {
    td.textContent = value || "—";
    td.className = value ? "" : "muted";
    return td;
  }
  const input = document.createElement("input");
  input.className = "cell";
  input.value = value ?? "";
  input.addEventListener("change", () => onChange(input.value.trim()));
  td.appendChild(input);
  return td;
}

/* ===================== EXCEL IMPORT ===================== */

async function onImportClients() {
  if (session.role !== "coordonator") return alert("Doar coordonatorul poate importa clienți.");

  const file = $("#fileClients").files?.[0];
  if (!file) return alert("Alege un fișier Excel.");

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const requestedSheet = ($("#sheetName").value || "").trim();
    const sheetName = requestedSheet && wb.SheetNames.includes(requestedSheet)
      ? requestedSheet
      : wb.SheetNames[0];

    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    if (!rows.length) return alert("Sheet-ul pare gol.");

    const keys = Object.keys(rows[0] || {});
    const nameKey = findKey(keys, ["client", "nume", "nume client", "denumire", "customer", "name"]);
    const countyKey = findKey(keys, ["judet", "județ", "county"]);
    const addrKey = findKey(keys, ["adresa", "adresă", "address", "locatie", "localizare"]);
    const phoneKey = findKey(keys, ["telefon", "phone", "tel"]);

    if (!nameKey) {
      return alert("Nu am găsit coloana de nume client. Folosește o coloană numită: Client / Nume / Nume client.");
    }

    // Build maps for dedupe
    const norm = s => String(s||"").trim().toLowerCase();
    const byNameCounty = new Map(); // "name|county" -> client
    const byNameOnly = new Map();   // "name" -> client

    for (const c of state.clients) {
      const nk = norm(c.name);
      const ck = norm(c.county);
      if (nk) {
        byNameOnly.set(nk, c);
        byNameCounty.set(`${nk}|${ck}`, c);
      }
    }

    let added = 0, updated = 0, skipped = 0;

    for (const r of rows) {
      const name = String(r[nameKey] ?? "").trim();
      if (!name) continue;

      const county = countyKey ? String(r[countyKey] ?? "").trim() : "";
      const address = addrKey ? String(r[addrKey] ?? "").trim() : "";
      const phone = phoneKey ? String(r[phoneKey] ?? "").trim() : "";

      const nk = norm(name);
      const ck = norm(county);

      // Prefer Name+County match; fallback to Name-only
      const existing = byNameCounty.get(`${nk}|${ck}`) || byNameOnly.get(nk);

      if (existing) {
        if (!existing.county && county) existing.county = county;
        if (!existing.address && address) existing.address = address;
        if (!existing.phone && phone) existing.phone = phone;
        updated++;
        // refresh maps if county was empty and now set
        byNameOnly.set(nk, existing);
        byNameCounty.set(`${nk}|${norm(existing.county)}`, existing);
      } else {
        const c = newClient({ name, county, address, phone });
        state.clients.push(c);
        byNameOnly.set(nk, c);
        byNameCounty.set(`${nk}|${ck}`, c);
        added++;
      }
    }

    alert(`Import gata.\nAdăugați: ${added}\nActualizați: ${updated}\nSheet: ${sheetName}`);
    $("#clientSearch").value = "";
    saveRender();
  } catch (e) {
    console.error(e);
    alert("Eroare la import. Verifică fișierul Excel.");
  }
}

function findKey(keys, candidates) {
  const norm = s => s.toLowerCase().replace(/\s+/g, " ").trim();
  const nkeys = keys.map(k => ({ raw: k, n: norm(k) }));
  for (const c of candidates) {
    const nc = norm(c);
    const hit = nkeys.find(k => k.n === nc || k.n.includes(nc));
    if (hit) return hit.raw;
  }
  return null;
}

/* ===================== AGENTS ===================== */

function onAddAgent() {
  if (session.role !== "coordonator") return;
  const name = $("#newAgentName").value.trim();
  if (!name) return alert("Completează numele agentului.");
  state.agents.push({ id: uid(), name });
  $("#newAgentName").value = "";
  saveRender();
}

function renderAgents() {
  const tbody = $("#tblAgents tbody");
  tbody.innerHTML = "";

  for (const a of (state.agents || []).slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const tr = document.createElement("tr");

    if (session.role === "coordonator") {
      tr.appendChild(tdInlineEdit(a.name, v => { a.name = v; saveRender(); }, true));
    } else {
      const td = document.createElement("td");
      td.textContent = a.name;
      tr.appendChild(td);
    }

    const tdDel = document.createElement("td");
    if (session.role === "coordonator") {
      const btn = document.createElement("button");
      btn.className = "btn danger";
      btn.textContent = "Șterge";
      btn.addEventListener("click", () => {
        if (!confirm(`Ștergi agentul "${a.name}"?`)) return;
        state.agents = state.agents.filter(x => x.id !== a.id);
        for (const r of state.routes) if (r.agentId === a.id) r.agentId = null;
        for (const v of state.visits) if (v.agentId === a.id) v.agentId = null;

        // also detach users linked to agent
        for (const u of state.users) if (u.agentId === a.id) u.agentId = null;

        saveRender();
      });
      tdDel.appendChild(btn);
    } else {
      tdDel.textContent = "—";
      tdDel.className = "muted";
    }
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }

  // update route agent selector
  const sel = $("#newRouteAgent");
  sel.innerHTML = `<option value="">(fără agent)</option>` + (state.agents || [])
    .slice().sort((x,y)=>x.name.localeCompare(y.name))
    .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");
}

/* ===================== ROUTES ===================== */

function onAddRoute() {
  if (session.role !== "coordonator") return;
  const name = $("#newRouteName").value.trim();
  const agentId = $("#newRouteAgent").value || null;
  if (!name) return alert("Completează numele rutei.");
  state.routes.push({ id: uid(), name, agentId });
  $("#newRouteName").value = "";
  saveRender();
}

function renderRoutes() {
  const tbody = $("#tblRoutes tbody");
  tbody.innerHTML = "";

  for (const r of (state.routes || []).slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const tr = document.createElement("tr");

    if (session.role === "coordonator") {
      tr.appendChild(tdInlineEdit(r.name, v => { r.name = v; saveRender(); }, true));
    } else {
      const td = document.createElement("td");
      td.textContent = r.name;
      tr.appendChild(td);
    }

    const tdAgent = document.createElement("td");
    if (session.role === "coordonator") {
      const sel = document.createElement("select");
      sel.className = "cell";
      sel.innerHTML = `<option value="">(fără)</option>` + (state.agents || [])
        .slice().sort((x,y)=>x.name.localeCompare(y.name))
        .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");
      sel.value = r.agentId || "";
      sel.addEventListener("change", () => {
        r.agentId = sel.value || null;
        saveRender();
      });
      tdAgent.appendChild(sel);
    } else {
      tdAgent.textContent = getAgentName(r.agentId) || "—";
      if (!r.agentId) tdAgent.className = "muted";
    }
    tr.appendChild(tdAgent);

    const tdDel = document.createElement("td");
    if (session.role === "coordonator") {
      const btn = document.createElement("button");
      btn.className = "btn danger";
      btn.textContent = "Șterge";
      btn.addEventListener("click", () => {
        if (!confirm(`Ștergi ruta "${r.name}"?`)) return;
        state.routes = state.routes.filter(x => x.id !== r.id);
        for (const v of state.visits) if (v.routeId === r.id) v.routeId = null;
        saveRender();
      });
      tdDel.appendChild(btn);
    } else {
      tdDel.textContent = "—";
      tdDel.className = "muted";
    }
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

/* ===================== ACTIVITIES ===================== */

function onAddActivity() {
  if (session.role !== "coordonator") return alert("Doar coordonatorul poate adăuga activități.");
  const name = $("#newActivityName").value.trim();
  if (!name) return alert("Completează activitatea.");
  state.activities.push({ id: uid(), name });
  $("#newActivityName").value = "";
  saveRender();
}

function renderActivities() {
  const tbody = $("#tblActivities tbody");
  tbody.innerHTML = "";

  for (const a of (state.activities || []).slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const tr = document.createElement("tr");
    if (session.role === "coordonator") {
      tr.appendChild(tdInlineEdit(a.name, v => { a.name = v; saveRender(); }, true));
    } else {
      const td = document.createElement("td");
      td.textContent = a.name;
      tr.appendChild(td);
    }

    const tdDel = document.createElement("td");
    if (session.role === "coordonator") {
      const btn = document.createElement("button");
      btn.className = "btn danger";
      btn.textContent = "Șterge";
      btn.addEventListener("click", () => {
        if (!confirm(`Ștergi activitatea "${a.name}"?`)) return;
        state.activities = state.activities.filter(x => x.id !== a.id);
        for (const v of state.visits) v.activityIds = (v.activityIds || []).filter(id => id !== a.id);
        saveRender();
      });
      tdDel.appendChild(btn);
    } else {
      tdDel.textContent = "—";
      tdDel.className = "muted";
    }
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }

  renderActivitiesChips($("#visitActivities"), [], () => {});
}

function renderActivitiesChips(container, selectedIds, onChange) {
  container.innerHTML = "";
  const set = new Set(selectedIds || []);
  for (const a of (state.activities || []).slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const chip = document.createElement("div");
    chip.className = "chip" + (set.has(a.id) ? " on" : "");
    chip.textContent = a.name;
    chip.addEventListener("click", () => {
      if (set.has(a.id)) set.delete(a.id); else set.add(a.id);
      onChange([...set]);
      renderActivitiesChips(container, [...set], onChange);
    });
    container.appendChild(chip);
  }
}

/* ===================== PLANNING SELECTORS ===================== */

function renderPlanningSelectors() {
  // Clients: show "Name (ABR)"
  const selC = $("#visitClient");
  selC.innerHTML = (state.clients || [])
    .slice().sort((a,b)=> (a.name||"").localeCompare(b.name||""))
    .map(c => `<option value="${c.id}">${escapeHtml(clientLabel(c))}</option>`)
    .join("");

  // Agents: if agent role, lock to own agent
  const selA = $("#visitAgent");
  if (session.role === "agent") {
    selA.innerHTML = `<option value="${escapeAttr(session.agentId||"")}">${escapeHtml(getAgentName(session.agentId)||"(agent)")}</option>`;
    selA.value = session.agentId || "";
    selA.disabled = true;
  } else {
    selA.disabled = false;
    selA.innerHTML = `<option value="">(fără)</option>` + (state.agents || [])
      .slice().sort((a,b)=>a.name.localeCompare(b.name))
      .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");
  }

  // Routes: agent sees all routes but when selecting route, agent auto-assign will happen
  const selR = $("#visitRoute");
  selR.innerHTML = `<option value="">(fără)</option>` + (state.routes || [])
    .slice().sort((a,b)=>a.name.localeCompare(b.name))
    .map(r => `<option value="${r.id}">${escapeHtml(r.name)}</option>`).join("");

  renderActivitiesChips($("#visitActivities"), [], () => {});
}

function renderAgentFilter() {
  const wrap = $("#agentFilterWrap");
  const sel = $("#agentFilter");

  if (session.role === "agent") {
    wrap.style.display = "none";
    sel.innerHTML = "";
    return;
  }

  wrap.style.display = "";
  const cur = sel.value || "ALL";
  sel.innerHTML = `<option value="ALL">Toți</option>` + (state.agents || [])
    .slice().sort((a,b)=>a.name.localeCompare(b.name))
    .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");
  // preserve selection if possible
  sel.value = (cur && [...sel.options].some(o => o.value === cur)) ? cur : "ALL";
}

function clientLabel(c) {
  const ab = countyAbbr(c.county || "");
  return ab ? `${c.name} (${ab})` : (c.name || "");
}

function countyAbbr(county) {
  const s = String(county || "").trim();
  if (!s) return "";
  // If user already typed short code like "MM", keep it
  if (/^[A-Z]{2,3}$/.test(s)) return s;

  // Create a compact abbr: first letters of up to 2 words
  const parts = s.toUpperCase().split(/\s+/).filter(Boolean);
  if (parts.length === 1) return parts[0].slice(0,2);
  return (parts[0][0] + parts[1][0]).slice(0,3);
}

/* ===================== VISITS + ROLE FILTERS ===================== */

function onGenerateVisits() {
  const isCoord = session.role === "coordonator";
  const { from, to } = getInterval();
  if (!from || !to) return alert("Alege intervalul (De la / Până la).");
  if (from > to) return alert("Interval invalid.");

  // agent role: generate only for their agent? In MVP: we still generate unassigned visits; but agent should only see their own.
  // We'll generate visits as unassigned; coordinator can assign. Agent should not generate bulk to avoid confusion.
  if (!isCoord) return alert("În MVP, doar coordonatorul generează automat vizite.");

  const existingKey = new Set(state.visits.map(v => `${v.clientId}|${v.date}`));
  let added = 0;

  const months = listMonthsInRange(from, to);
  for (const c of (state.clients || [])) {
    const weeks = (c.monthlyWeeks || []).slice().sort((a,b)=>a-b);
    const wantCount = clampInt(c.monthlyCount ?? weeks.length, 0, 50);

    for (const { y, m } of months) {
      const datesForWeeks = weeks
        .map(w => firstDateOfWeekOfMonth(y, m, w))
        .filter(d => d && d >= from && d <= to)
        .map(d => toISO(d));

      const chosen = datesForWeeks.slice(0, wantCount);
      for (const iso of chosen) {
        const k = `${c.id}|${iso}`;
        if (existingKey.has(k)) continue;
        state.visits.push({
          id: uid(),
          date: iso,
          clientId: c.id,
          agentId: null,
          routeId: null,
          activityIds: [],
          otherActivity: "",
          details: "",
          obs: ""
        });
        existingKey.add(k);
        added++;
      }
    }
  }

  alert(`Generare gata. Vizite adăugate: ${added}`);
  saveRender();
}

function onClearVisitsInInterval() {
  const isCoord = session.role === "coordonator";
  if (!isCoord) return alert("În MVP, doar coordonatorul poate șterge vizite în masă.");
  const { from, to } = getInterval();
  if (!from || !to) return alert("Alege intervalul.");
  if (!confirm("Ștergi toate vizitele din interval?")) return;

  state.visits = state.visits.filter(v => {
    const d = fromISO(v.date);
    return !(d >= from && d <= to);
  });
  saveRender();
}

function onAddVisitManual() {
  const date = $("#visitDate").value;
  const clientId = $("#visitClient").value;
  const agentIdUI = $("#visitAgent").value || null;
  const routeId = $("#visitRoute").value || null;

  if (!date) return alert("Alege data.");
  if (!clientId) return alert("Alege client.");

  // agent role: force own agent
  const agentId = (session.role === "agent") ? (session.agentId || null) : agentIdUI;

  // selected activities:
  const selectedActivityIds = [...$("#visitActivities").querySelectorAll(".chip.on")]
    .map(ch => {
      const name = ch.textContent;
      const a = state.activities.find(x => x.name === name);
      return a?.id;
    })
    .filter(Boolean);

  const otherActivity = $("#visitOtherActivity").value.trim();
  const details = $("#visitDetails").value.trim();
  const obs = $("#visitObs").value.trim();

  // If route has agent and agent not set, auto set (coord only)
  let finalAgent = agentId;
  if (routeId && !finalAgent && session.role === "coordonator") {
    const r = state.routes.find(x => x.id === routeId);
    if (r?.agentId) finalAgent = r.agentId;
  }

  state.visits.push({
    id: uid(),
    date,
    clientId,
    agentId: finalAgent,
    routeId,
    activityIds: selectedActivityIds,
    otherActivity,
    details,
    obs
  });

  $("#visitOtherActivity").value = "";
  $("#visitDetails").value = "";
  $("#visitObs").value = "";

  saveRender();
}

function renderVisits() {
  const tbody = $("#tblVisits tbody");
  tbody.innerHTML = "";

  const { from, to } = getInterval();
  const inInterval = (v) => {
    if (!from || !to) return true;
    const d = fromISO(v.date);
    return d >= from && d <= to;
  };

  const agentFilter = (session.role === "coordonator") ? ($("#agentFilter").value || "ALL") : "ALL";

  const canSee = (v) => {
    if (session.role === "agent") return v.agentId && v.agentId === session.agentId;
    if (agentFilter !== "ALL") return v.agentId === agentFilter;
    return true;
  };

  const visits = (state.visits || []).filter(inInterval).filter(canSee);

  for (const v of visits) {
    const tr = document.createElement("tr");

    tr.appendChild(tdText(dayFromISO(v.date)));
    tr.appendChild(tdDate(v.date, newDate => {
      if (!canEditVisit(v)) return;
      v.date = newDate;
      saveRender();
    }, canEditVisit(v)));

    // client select (both can edit if own visit)
    tr.appendChild(tdSelect(
      (state.clients || []).slice().sort((a,b)=> (a.name||"").localeCompare(b.name||"")),
      v.clientId,
      c => c.id,
      c => clientLabel(c),
      newId => { if (canEditVisit(v)) { v.clientId = newId; saveRender(); } },
      canEditVisit(v)
    ));

    // agent select (agent locked)
    tr.appendChild(tdSelectAgentForVisit(v));

    // route select
    tr.appendChild(tdSelect(
      [{id:"", name:"(fără)"}, ...(state.routes || []).slice().sort((a,b)=>a.name.localeCompare(b.name))],
      v.routeId || "",
      r => r.id,
      r => r.name,
      newId => {
        if (!canEditVisit(v)) return;
        v.routeId = newId || null;
        // if route has agent and visit agent empty, auto set for coord
        if (session.role === "coordonator") {
          const r = state.routes.find(x => x.id === v.routeId);
          if (r?.agentId && !v.agentId) v.agentId = r.agentId;
        }
        saveRender();
      },
      canEditVisit(v)
    ));

    // activities editor
    const tdAct = document.createElement("td");
    tdAct.appendChild(activityEditor(v, () => saveRender(), canEditVisit(v)));
    tr.appendChild(tdAct);

    tr.appendChild(tdInput(v.details || "", val => { if (canEditVisit(v)) { v.details = val; saveRender(); } }, canEditVisit(v)));
    tr.appendChild(tdInput(v.obs || "", val => { if (canEditVisit(v)) { v.obs = val; saveRender(); } }, canEditVisit(v)));

    const tdDel = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "btn danger";
    btn.textContent = "Șterge";
    btn.disabled = !canEditVisit(v);
    btn.addEventListener("click", () => {
      if (!canEditVisit(v)) return;
      if (!confirm("Ștergi vizita?")) return;
      state.visits = state.visits.filter(x => x.id !== v.id);
      saveRender();
    });
    tdDel.appendChild(btn);
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

function canEditVisit(v) {
  if (session.role === "coordonator") return true;
  return session.role === "agent" && v.agentId === session.agentId;
}

function tdSelectAgentForVisit(v) {
  const td = document.createElement("td");

  if (session.role === "agent") {
    const sel = document.createElement("select");
    sel.className = "cell";
    sel.innerHTML = `<option value="${escapeAttr(session.agentId||"")}">${escapeHtml(getAgentName(session.agentId)||"(agent)")}</option>`;
    sel.value = session.agentId || "";
    sel.disabled = true;
    td.appendChild(sel);
    // enforce agent id
    if (!v.agentId) v.agentId = session.agentId || null;
    return td;
  }

  // coord
  const items = [{id:"", name:"(fără)"}, ...(state.agents || []).slice().sort((a,b)=>a.name.localeCompare(b.name))];
  const sel = document.createElement("select");
  sel.className = "cell";
  sel.innerHTML = items.map(a => `<option value="${escapeAttr(String(a.id))}">${escapeHtml(String(a.name))}</option>`).join("");
  sel.value = v.agentId || "";
  sel.addEventListener("change", () => {
    if (!canEditVisit(v)) return;
    v.agentId = sel.value || null;
    saveRender();
  });
  td.appendChild(sel);
  return td;
}

function activityEditor(visit, onSave, enabled=true) {
  const wrap = document.createElement("div");
  wrap.style.display = "grid";
  wrap.style.gap = "8px";

  const box = document.createElement("div");
  box.className = "multibox";
  const label = document.createElement("div");
  label.className = "label";
  label.textContent = "Alege (multiple)";
  box.appendChild(label);

  const chips = document.createElement("div");
  chips.className = "chips";

  const set = new Set(visit.activityIds || []);
  for (const a of (state.activities || []).slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const chip = document.createElement("div");
    chip.className = "chip" + (set.has(a.id) ? " on" : "");
    chip.textContent = a.name;
    chip.style.opacity = enabled ? "1" : "0.65";
    chip.style.pointerEvents = enabled ? "auto" : "none";
    chip.addEventListener("click", () => {
      if (!enabled) return;
      if (set.has(a.id)) set.delete(a.id); else set.add(a.id);
      visit.activityIds = [...set];
      onSave();
    });
    chips.appendChild(chip);
  }
  box.appendChild(chips);

  const other = document.createElement("input");
  other.className = "cell";
  other.placeholder = "Altceva (opțional)";
  other.value = visit.otherActivity || "";
  other.disabled = !enabled;
  other.addEventListener("change", () => {
    if (!enabled) return;
    visit.otherActivity = other.value.trim();
    onSave();
  });

  wrap.appendChild(box);
  wrap.appendChild(other);
  return wrap;
}

function onSortVisits() {
  if (!session) return;
  const mode = $("#sortMode").value;

  const clientsById = new Map((state.clients||[]).map(c => [c.id, c.name]));
  const agentsById  = new Map((state.agents||[]).map(a => [a.id, a.name]));
  const routesById  = new Map((state.routes||[]).map(r => [r.id, r.name]));

  const key = (v) => {
    if (mode === "client") return (clientsById.get(v.clientId) || "");
    if (mode === "agent") return (agentsById.get(v.agentId) || "");
    if (mode === "route") return (routesById.get(v.routeId) || "");
    return v.date || "";
  };

  state.visits.sort((a,b) => {
    const ka = key(a), kb = key(b);
    if (mode === "date") return (a.date || "").localeCompare(b.date || "");
    const cmp = String(ka).localeCompare(String(kb));
    if (cmp !== 0) return cmp;
    return (a.date || "").localeCompare(b.date || "");
  });

  saveRender();
}

/* ===================== EXPORT ===================== */

function onExport() {
  if (!session) return;

  const { from, to } = getInterval();
  const visits = (state.visits || []).filter(v => {
    if (!from || !to) return true;
    const d = fromISO(v.date);
    return d >= from && d <= to;
  });

  // role filter
  const agentFilter = (session.role === "coordonator") ? ($("#agentFilter").value || "ALL") : "ALL";
  const filtered = visits.filter(v => {
    if (session.role === "agent") return v.agentId === session.agentId;
    if (agentFilter !== "ALL") return v.agentId === agentFilter;
    return true;
  });

  const clientsById = new Map((state.clients||[]).map(c => [c.id, c.name]));
  const actsById = new Map((state.activities||[]).map(a => [a.id, a.name]));

  const rows = [
    ["Zi", "Data", "Client", "Activitate", "Detalii", "Observatii"]
  ];

  const sorted = filtered.slice().sort((a,b)=> (a.date||"").localeCompare(b.date||""));
  for (const v of sorted) {
    const zi = dayFromISO(v.date);
    const data = formatDateRO(v.date); // DD.MM.YYYY
    const client = clientsById.get(v.clientId) || "";
    const actNames = (v.activityIds || []).map(id => actsById.get(id)).filter(Boolean);
    let act = actNames.join(", ");
    if (v.otherActivity && v.otherActivity.trim()) {
      act = act ? `${act}, ${v.otherActivity.trim()}` : v.otherActivity.trim();
    }
    rows.push([zi, data, client, act, v.details || "", v.obs || ""]);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [
    { wch: 10 }, { wch: 12 }, { wch: 28 },
    { wch: 28 }, { wch: 22 }, { wch: 22 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Planificare");
  const name = buildExportName(from, to);
  XLSX.writeFile(wb, name);
}

function buildExportName(from, to) {
  const f = from ? toISO(from) : "ALL";
  const t = to ? toISO(to) : "ALL";
  return `planificare_${f}_-${t}.xlsx`;
}

/* ===================== REPORTS ===================== */

function renderReport() {
  const { from, to } = getInterval();
  const visits = (state.visits || []).filter(v => {
    if (!from || !to) return true;
    const d = fromISO(v.date);
    return d >= from && d <= to;
  });

  // role filter
  const agentFilter = (session.role === "coordonator") ? ($("#agentFilter").value || "ALL") : "ALL";
  const filtered = visits.filter(v => {
    if (session.role === "agent") return v.agentId === session.agentId;
    if (agentFilter !== "ALL") return v.agentId === agentFilter;
    return true;
  });

  const uniqueClients = new Set(filtered.map(v => v.clientId).filter(Boolean));

  const byAgent = new Map();
  const byRoute = new Map();
  for (const v of filtered) {
    const a = v.agentId || "(fără)";
    const r = v.routeId || "(fără)";
    byAgent.set(a, (byAgent.get(a) || 0) + 1);
    byRoute.set(r, (byRoute.get(r) || 0) + 1);
  }

  const agentsById = new Map((state.agents||[]).map(a => [a.id, a.name]));
  const routesById = new Map((state.routes||[]).map(r => [r.id, r.name]));

  const stats = $("#reportStats");
  stats.innerHTML = "";

  stats.appendChild(statCard("Nr. vizite (interval)", String(filtered.length)));
  stats.appendChild(statCard("Nr. clienți unici (interval)", String(uniqueClients.size)));

  stats.appendChild(statCard("Vizite pe agent (top)", topList(byAgent, id => agentsById.get(id) || id)));
  stats.appendChild(statCard("Vizite pe rută (top)", topList(byRoute, id => routesById.get(id) || id)));
}

function statCard(k, v) {
  const d = document.createElement("div");
  d.className = "stat";
  const kk = document.createElement("div");
  kk.className = "k";
  kk.textContent = k;
  const vv = document.createElement("div");
  vv.className = "v";
  vv.innerHTML = v;
  d.appendChild(kk);
  d.appendChild(vv);
  return d;
}

function topList(map, labelFn) {
  const arr = [...map.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 8);
  if (!arr.length) return "<span style='color:#7f8db3'>—</span>";
  return arr.map(([id, n]) => `${escapeHtml(labelFn(id))}: <b>${n}</b>`).join("<br/>");
}

/* ===================== USERS (COORD) ===================== */

function onCreateUser() {
  if (session.role !== "coordonator") return;

  const username = $("#newUsername").value.trim();
  const password = $("#newPassword").value;
  const role = $("#newRole").value;
  const agentId = $("#newUserAgent").value || null;

  if (!username) return alert("Completează username.");
  if (!password || password.length < 4) return alert("Parola trebuie să aibă minim 4 caractere.");
  if (state.users.some(u => u.username === username)) return alert("Username deja existent.");

  if (role === "agent" && !agentId) return alert("Pentru agent, alege agentul asociat.");

  state.users.push({
    id: uid(),
    username,
    passHash: hash(password),
    role,
    agentId: role === "agent" ? agentId : null
  });

  $("#newUsername").value = "";
  $("#newPassword").value = "";
  $("#newRole").value = "agent";
  $("#newUserAgent").value = "";

  saveRender();
}

function renderUsers() {
  if (!session) return;

  // populate agent selector for user creation
  const sel = $("#newUserAgent");
  sel.innerHTML = `<option value="">(alege agent)</option>` + (state.agents || [])
    .slice().sort((a,b)=>a.name.localeCompare(b.name))
    .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");

  const panel = $("#tab-users");
  if (session.role !== "coordonator") {
    panel.innerHTML = `<div class="card"><h2>Acces restricționat</h2><div class="hint">Doar coordonatorul.</div></div>`;
    return;
  }

  const tbody = $("#tblUsers tbody");
  tbody.innerHTML = "";

  for (const u of (state.users || []).slice().sort((a,b)=>a.username.localeCompare(b.username))) {
    const tr = document.createElement("tr");
    tr.appendChild(tdText(u.username));
    tr.appendChild(tdText(u.role));
    tr.appendChild(tdText(u.role === "agent" ? (getAgentName(u.agentId) || "—") : "—"));

    const tdDel = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "btn danger";
    btn.textContent = "Șterge";
    btn.addEventListener("click", () => {
      if (u.username === "admin") return alert("Admin nu poate fi șters în MVP.");
      if (!confirm(`Ștergi user "${u.username}"?`)) return;
      state.users = state.users.filter(x => x.id !== u.id);
      saveRender();
    });
    tdDel.appendChild(btn);
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

/* ===================== CLIENT MODAL (FIȘĂ CLIENT) ===================== */

let modalClientId = null;

function openClientModal(clientId) {
  modalClientId = clientId;
  const c = state.clients.find(x => x.id === clientId);
  if (!c) return;

  $("#cmTitle").textContent = `Fișă Client: ${c.name || ""}`;
  $("#cmSubtitle").textContent = `Județ: ${c.county || "—"} • Telefon: ${c.phone || "—"}`;

  $("#cmName").value = c.name || "";
  $("#cmCounty").value = c.county || "";
  $("#cmAddress").value = c.address || "";
  $("#cmPhone").value = c.phone || "";

  $("#cmCount").textContent = String(c.monthlyCount ?? 0);

  // weeks chips
  renderWeeksInModal(c);

  // contacts
  renderContactsInModal(c);

  // role: only coordinator can edit + delete
  const canEdit = session.role === "coordonator";
  $("#cmName").disabled = !canEdit;
  $("#cmCounty").disabled = !canEdit;
  $("#cmAddress").disabled = !canEdit;
  $("#cmPhone").disabled = !canEdit;
  $("#cmMinus").disabled = !canEdit;
  $("#cmPlus").disabled = !canEdit;
  $("#btnSaveClientModal").style.display = canEdit ? "" : "none";
  $("#btnDeleteClientModal").style.display = canEdit ? "" : "none";
  $("#btnAddContact").style.display = canEdit ? "" : "none";

  $("#clientModal").classList.remove("hidden");
}

function closeClientModal() {
  modalClientId = null;
  $("#clientModal").classList.add("hidden");
}

function changeClientModalCount(delta) {
  if (session.role !== "coordonator") return;
  const c = state.clients.find(x => x.id === modalClientId);
  if (!c) return;
  c.monthlyCount = clampInt((c.monthlyCount ?? 0) + delta, 0, 50);
  $("#cmCount").textContent = String(c.monthlyCount);
  saveRender(false); // don't close modal
  // keep modal updated
}

function renderWeeksInModal(c) {
  const wrap = $("#cmWeeks");
  wrap.innerHTML = "";
  const selected = new Set(c.monthlyWeeks || []);
  for (let w = 1; w <= 5; w++) {
    const chip = document.createElement("div");
    chip.className = "chip" + (selected.has(w) ? " on" : "");
    chip.textContent = `S${w}`;
    chip.style.opacity = session.role === "coordonator" ? "1" : "0.65";
    chip.style.pointerEvents = session.role === "coordonator" ? "auto" : "none";
    chip.addEventListener("click", () => {
      if (session.role !== "coordonator") return;
      if (selected.has(w)) selected.delete(w); else selected.add(w);
      c.monthlyWeeks = [...selected].sort((a,b)=>a-b);
      saveRender(false);
      renderWeeksInModal(c);
    });
    wrap.appendChild(chip);
  }
}

function renderContactsInModal(c) {
  if (!Array.isArray(c.contacts)) c.contacts = [];

  const wrap = $("#cmContacts");
  wrap.innerHTML = "";

  c.contacts.forEach((ct, idx) => {
    const card = document.createElement("div");
    card.className = "contactCard";

    const top = document.createElement("div");
    top.className = "contactTop";

    const name = document.createElement("div");
    name.className = "name";
    name.textContent = ct.name ? ct.name : `Contact #${idx+1}`;

    const del = document.createElement("button");
    del.className = "btn danger";
    del.textContent = "Șterge";
    del.style.padding = "8px 10px";
    del.disabled = session.role !== "coordonator";
    del.addEventListener("click", () => {
      if (session.role !== "coordonator") return;
      if (!confirm("Ștergi contactul?")) return;
      c.contacts.splice(idx, 1);
      saveRender(false);
      renderContactsInModal(c);
    });

    top.appendChild(name);
    top.appendChild(del);
    card.appendChild(top);

    const grid = document.createElement("div");
    grid.className = "grid2";
    grid.style.marginTop = "10px";
    grid.style.gap = "10px";

    grid.appendChild(fieldInput("Nume", ct.name || "", v => { ct.name = v; saveRender(false); }, session.role === "coordonator"));
    grid.appendChild(fieldInput("Rol/Funcție", ct.role || "", v => { ct.role = v; saveRender(false); }, session.role === "coordonator"));
    grid.appendChild(fieldInput("Telefon", ct.phone || "", v => { ct.phone = v; saveRender(false); }, session.role === "coordonator"));
    grid.appendChild(fieldInput("Email", ct.email || "", v => { ct.email = v; saveRender(false); }, session.role === "coordonator"));

    const obs = document.createElement("div");
    obs.className = "field";
    const lbl = document.createElement("span");
    lbl.textContent = "Observații";
    const input = document.createElement("input");
    input.value = ct.notes || "";
    input.disabled = session.role !== "coordonator";
    input.addEventListener("change", () => { ct.notes = input.value.trim(); saveRender(false); });
    obs.appendChild(lbl);
    obs.appendChild(input);

    card.appendChild(grid);
    card.appendChild(obs);

    wrap.appendChild(card);
  });
}

function fieldInput(label, value, onChange, enabled=true) {
  const f = document.createElement("label");
  f.className = "field";
  const s = document.createElement("span");
  s.textContent = label;
  const i = document.createElement("input");
  i.value = value || "";
  i.disabled = !enabled;
  i.addEventListener("change", () => onChange(i.value.trim()));
  f.appendChild(s);
  f.appendChild(i);
  return f;
}

function addContactInModal() {
  if (session.role !== "coordonator") return;
  const c = state.clients.find(x => x.id === modalClientId);
  if (!c) return;
  if (!Array.isArray(c.contacts)) c.contacts = [];
  c.contacts.push({ id: uid(), name:"", role:"", phone:"", email:"", notes:"" });
  saveRender(false);
  renderContactsInModal(c);
}

function saveClientModal() {
  if (session.role !== "coordonator") return;
  const c = state.clients.find(x => x.id === modalClientId);
  if (!c) return;

  c.name = $("#cmName").value.trim();
  c.county = $("#cmCounty").value.trim();
  c.address = $("#cmAddress").value.trim();
  c.phone = $("#cmPhone").value.trim();

  saveRender(false);
  // refresh subtitle
  $("#cmSubtitle").textContent = `Județ: ${c.county || "—"} • Telefon: ${c.phone || "—"}`;
  alert("Salvat ✅");
}

function deleteClientFromModal() {
  if (session.role !== "coordonator") return;
  const c = state.clients.find(x => x.id === modalClientId);
  if (!c) return;
  if (!confirm(`Ștergi clientul "${c.name}"?`)) return;
  state.clients = state.clients.filter(x => x.id !== c.id);
  state.visits = state.visits.filter(v => v.clientId !== c.id);
  closeClientModal();
  saveRender();
}

/* ===================== RESET ===================== */

function onResetAll() {
  if (session.role !== "coordonator") return;
  if (!confirm("Reset total? Se șterg clienți, agenți, rute, activități, vizite și utilizatori (mai puțin admin).")) return;

  const admin = (state.users || []).find(u => u.username === "admin") || null;

  state = {
    clients: [],
    agents: [],
    routes: [],
    activities: [
      { id: uid(), name: "comanda" },
      { id: uid(), name: "incasare" },
      { id: uid(), name: "comanda + incasare" },
    ],
    visits: [],
    users: admin ? [admin] : []
  };
  saveRender();
}

/* ===================== HELPERS UI ===================== */

function saveRender(closeModal=true) {
  saveState(state);
  // keep modal open if asked
  renderAll();
  if (!closeModal && modalClientId) {
    $("#clientModal").classList.remove("hidden");
  }
}

function tdInput(value, onChange, enabled=true) {
  const td = document.createElement("td");
  const input = document.createElement("input");
  input.className = "cell";
  input.value = value ?? "";
  input.disabled = !enabled;
  input.addEventListener("change", () => onChange(input.value.trim()));
  td.appendChild(input);
  return td;
}

function tdText(value) {
  const td = document.createElement("td");
  td.textContent = value ?? "";
  return td;
}

function tdDate(iso, onChange, enabled=true) {
  const td = document.createElement("td");
  const input = document.createElement("input");
  input.type = "date";
  input.className = "cell";
  input.value = iso || "";
  input.disabled = !enabled;
  input.addEventListener("change", () => onChange(input.value));
  td.appendChild(input);
  return td;
}

function tdSelect(items, selectedValue, getVal, getLabel, onChange, enabled=true) {
  const td = document.createElement("td");
  const sel = document.createElement("select");
  sel.className = "cell";
  sel.disabled = !enabled;
  sel.innerHTML = items.map(it => {
    const v = getVal(it);
    const label = getLabel(it);
    return `<option value="${escapeAttr(String(v))}">${escapeHtml(String(label))}</option>`;
  }).join("");
  sel.value = selectedValue ?? "";
  sel.addEventListener("change", () => onChange(sel.value));
  td.appendChild(sel);
  return td;
}

function $(q) { return document.querySelector(q); }

/* ===================== DATE HELPERS ===================== */

function toISO(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${da}`;
}
function fromISO(iso) {
  const [y,m,d] = iso.split("-").map(Number);
  return new Date(y, m-1, d);
}
function dayFromISO(iso) {
  if (!iso) return "";
  const d = fromISO(iso);
  return dayNames[d.getDay()];
}
function formatDateRO(iso) {
  if (!iso) return "";
  const [y,m,d] = iso.split("-");
  return `${d}.${m}.${y}`;
}
function getInterval() {
  const f = $("#fromDate").value;
  const t = $("#toDate").value;
  return {
    from: f ? fromISO(f) : null,
    to: t ? fromISO(t) : null
  };
}

function listMonthsInRange(from, to) {
  const out = [];
  let y = from.getFullYear();
  let m = from.getMonth();
  const endY = to.getFullYear();
  const endM = to.getMonth();
  while (y < endY || (y === endY && m <= endM)) {
    out.push({ y, m });
    m++;
    if (m > 11) { m = 0; y++; }
  }
  return out;
}

// Week-of-month block (1..5): pick start day (1,8,15,22,29) then move to Monday within that block when possible
function firstDateOfWeekOfMonth(year, month0, weekNum) {
  const startDay = (weekNum - 1) * 7 + 1;
  const d = new Date(year, month0, startDay);
  if (d.getMonth() !== month0) return null;

  const day = d.getDay();
  let deltaToMon = (1 - day + 7) % 7;
  const candidate = new Date(d);
  candidate.setDate(d.getDate() + deltaToMon);

  if (candidate.getMonth() !== month0) return d;
  const candDay = candidate.getDate();
  if (candDay >= startDay && candDay <= startDay + 6) return candidate;
  return d;
}

/* ===================== PERSISTENCE ===================== */

function loadState() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return { clients: [], agents: [], routes: [], activities: [], visits: [], users: [] };
    const parsed = JSON.parse(raw);

    const s = {
      clients: Array.isArray(parsed.clients) ? parsed.clients : [],
      agents: Array.isArray(parsed.agents) ? parsed.agents : [],
      routes: Array.isArray(parsed.routes) ? parsed.routes : [],
      activities: Array.isArray(parsed.activities) ? parsed.activities : [],
      visits: Array.isArray(parsed.visits) ? parsed.visits : [],
      users: Array.isArray(parsed.users) ? parsed.users : [],
    };

    // migrations defaults
    for (const c of s.clients) {
      if (c.county === undefined) c.county = "";
      if (c.address === undefined) c.address = "";
      if (c.phone === undefined) c.phone = "";
      if (c.contacts === undefined) c.contacts = [];
      if (c.monthlyCount === undefined) c.monthlyCount = 2;
      if (!Array.isArray(c.monthlyWeeks)) c.monthlyWeeks = [1,3];
      if (!c.freqType) c.freqType = "monthly";
      if (c.reports === undefined) c.reports = {};
    }
    for (const u of s.users) {
      if (!u.passHash && u.password) { // legacy
        u.passHash = hash(u.password);
        delete u.password;
      }
    }

    return s;
  } catch {
    return { clients: [], agents: [], routes: [], activities: [], visits: [], users: [] };
  }
}

function saveState(s) {
  localStorage.setItem(LS_KEY, JSON.stringify(s));
}

function loadSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}
function saveSession(sess) {
  if (!sess) localStorage.removeItem(SESSION_KEY);
  else localStorage.setItem(SESSION_KEY, JSON.stringify(sess));
}

/* ===================== UTILS ===================== */

function uid() {
  return Math.random().toString(16).slice(2) + "-" + Date.now().toString(16);
}
function clampInt(v, min, max) {
  const n = parseInt(v, 10);
  if (Number.isNaN(n)) return min;
  return Math.max(min, Math.min(max, n));
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[c]));
}
function escapeAttr(s) {
  return String(s).replace(/"/g, "&quot;");
}

function getAgentName(agentId) {
  if (!agentId) return "";
  const a = (state.agents || []).find(x => x.id === agentId);
  return a ? a.name : "";
}

/* Tiny hash for local-only auth (NOT secure like server-side).
   Enough to avoid storing plain text passwords. */
function hash(str) {
  let h = 2166136261;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16);
}
