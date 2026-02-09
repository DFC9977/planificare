/* Planificare Vizite MVP
   - Import clienți din Excel
   - Frecvență lunară: săptămâni 1-5 selectate
   - Agenți & rute
   - Activități: listă configurabilă + multi-select + "altceva"
   - Planificare manual ajustabilă
   - Export Excel: Zi | Data | Client | Activitate | Detalii | Observatii
   - Persistență localStorage
*/

const LS_KEY = "pv_state_v1";

const dayNames = ["DUMINICA","LUNI","MARTI","MIERCURI","JOI","VINERI","SAMBATA"]; // fără diacritice

/** @type {State} */
let state = loadState();

/** Types (informal)
 * State: { clients: Client[], agents: Agent[], routes: Route[], activities: Activity[], visits: Visit[] }
 * Client: { id, name, address, freqType:"monthly", monthlyWeeks:number[], monthlyCount:number }
 * Agent: { id, name }
 * Route: { id, name, agentId|null }
 * Activity: { id, name }
 * Visit: { id, date:"YYYY-MM-DD", clientId, agentId|null, routeId|null, activityIds:string[], otherActivity:"", details:"", obs:"" }
 */

document.addEventListener("DOMContentLoaded", () => {
  bindTabs();

  // Buttons
  $("#btnReset").addEventListener("click", onResetAll);
  $("#btnExport").addEventListener("click", onExport);

  // Clients
  $("#btnImportClients").addEventListener("click", onImportClients);
  $("#btnAddClient").addEventListener("click", onAddClient);
  $("#btnClearClients").addEventListener("click", () => {
    if (!confirm("Ștergi toți clienții?")) return;
    state.clients = [];
    // păstrăm vizitele (poate vrei), dar curățăm referințe invalide
    state.visits = state.visits.filter(v => state.clients.some(c => c.id === v.clientId));
    saveRender();
  });
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

  // Reports
  $("#btnRefreshReport").addEventListener("click", renderReport);

  // Defaults
  if (!state.activities.length) {
    state.activities = [
      { id: uid(), name: "comanda" },
      { id: uid(), name: "incasare" },
      { id: uid(), name: "comanda + incasare" },
    ];
  }

  // initial dates: current month range
  const now = new Date();
  const from = new Date(now.getFullYear(), now.getMonth(), 1);
  const to = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  $("#fromDate").value = toISO(from);
  $("#toDate").value = toISO(to);
  $("#visitDate").value = toISO(now);

  saveRender();
});

function saveRender() {
  saveState(state);
  renderAll();
}

function renderAll() {
  renderClients();
  renderAgents();
  renderRoutes();
  renderActivities();
  renderPlanningSelectors();
  renderVisits();
  renderReport();
}

/* -------------------- Tabs -------------------- */
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

/* -------------------- Clients -------------------- */
function onAddClient() {
  const name = $("#newClientName").value.trim();
  const address = $("#newClientAddress").value.trim();
  if (!name) return alert("Completează numele clientului.");
  state.clients.push({
    id: uid(),
    name,
    address,
    freqType: "monthly",
    monthlyCount: 2,         // default
    monthlyWeeks: [1, 3],    // default bilunar 1&3
  });
  $("#newClientName").value = "";
  $("#newClientAddress").value = "";
  saveRender();
}

function renderClients() {
  const q = ($("#clientSearch").value || "").trim().toLowerCase();
  const tbody = $("#tblClients tbody");
  tbody.innerHTML = "";

  const filtered = state.clients
    .filter(c => !q || c.name.toLowerCase().includes(q) || (c.address || "").toLowerCase().includes(q))
    .sort((a,b)=> a.name.localeCompare(b.name));

  for (const c of filtered) {
    const tr = document.createElement("tr");

    // name
    tr.appendChild(tdInput(c.name, v => { c.name = v; saveRender(); }));

    // address
    tr.appendChild(tdInput(c.address || "", v => { c.address = v; saveRender(); }));

    // frequency (monthly count)
    const tdFreq = document.createElement("td");
    tdFreq.innerHTML = `
      <div class="row wrap">
        <select class="cell" data-id="${c.id}">
          <option value="monthly" selected>Lunar</option>
        </select>
        <input class="cell" type="number" min="0" step="1" value="${c.monthlyCount ?? 0}" style="max-width:110px" title="nr vizite/lună" />
      </div>
      <div class="hint">nr vizite/lună</div>
    `;
    const cntInput = tdFreq.querySelector("input");
    cntInput.addEventListener("change", () => {
      c.monthlyCount = clampInt(cntInput.value, 0, 50);
      saveRender();
    });
    tr.appendChild(tdFreq);

    // weeks 1-5
    const tdWeeks = document.createElement("td");
    tdWeeks.appendChild(weeksPicker(c.monthlyWeeks || [], weeks => {
      c.monthlyWeeks = weeks;
      saveRender();
    }));
    tr.appendChild(tdWeeks);

    // delete
    const tdDel = document.createElement("td");
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
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

function weeksPicker(selected, onChange) {
  const wrap = document.createElement("div");
  wrap.className = "row wrap";
  for (let w = 1; w <= 5; w++) {
    const chip = document.createElement("div");
    chip.className = "chip" + (selected.includes(w) ? " on" : "");
    chip.textContent = `S${w}`;
    chip.addEventListener("click", () => {
      const set = new Set(selected);
      if (set.has(w)) set.delete(w); else set.add(w);
      const arr = [...set].sort((a,b)=>a-b);
      onChange(arr);
    });
    wrap.appendChild(chip);
  }
  return wrap;
}

/* -------------------- Import Excel Clients -------------------- */
async function onImportClients() {
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

    // Detect columns
    const keys = Object.keys(rows[0] || {});
    const nameKey = findKey(keys, ["client", "nume", "nume client", "denumire", "customer", "name"]);
    const addrKey = findKey(keys, ["adresa", "adresă", "address", "locatie", "localizare"]);

    if (!nameKey) {
      return alert("Nu am găsit coloana de nume client. Folosește o coloană numită: Client / Nume / Nume client.");
    }

    // merge by name (case-insensitive)
    const existingByLower = new Map(state.clients.map(c => [c.name.trim().toLowerCase(), c]));
    let added = 0, updated = 0;

    for (const r of rows) {
      const name = String(r[nameKey] ?? "").trim();
      if (!name) continue;
      const address = addrKey ? String(r[addrKey] ?? "").trim() : "";

      const key = name.toLowerCase();
      const existing = existingByLower.get(key);
      if (existing) {
        // update address if empty
        if (!existing.address && address) existing.address = address;
        updated++;
      } else {
        state.clients.push({
          id: uid(),
          name,
          address,
          freqType: "monthly",
          monthlyCount: 2,
          monthlyWeeks: [1, 3],
        });
        added++;
      }
    }

    alert(`Import gata.\nAdăugați: ${added}\nActualizați: ${updated}\nSheet: ${sheetName}`);
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

/* -------------------- Agents -------------------- */
function onAddAgent() {
  const name = $("#newAgentName").value.trim();
  if (!name) return alert("Completează numele agentului.");
  state.agents.push({ id: uid(), name });
  $("#newAgentName").value = "";
  saveRender();
}

function renderAgents() {
  const tbody = $("#tblAgents tbody");
  tbody.innerHTML = "";
  for (const a of state.agents.slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const tr = document.createElement("tr");
    tr.appendChild(tdInput(a.name, v => { a.name = v; saveRender(); }));

    const tdDel = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "btn danger";
    btn.textContent = "Șterge";
    btn.addEventListener("click", () => {
      if (!confirm(`Ștergi agentul "${a.name}"?`)) return;
      state.agents = state.agents.filter(x => x.id !== a.id);
      // clear agentId references
      for (const r of state.routes) if (r.agentId === a.id) r.agentId = null;
      for (const v of state.visits) if (v.agentId === a.id) v.agentId = null;
      saveRender();
    });
    tdDel.appendChild(btn);
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }

  // update route agent selector
  const sel = $("#newRouteAgent");
  sel.innerHTML = `<option value="">(fără agent)</option>` + state.agents
    .slice().sort((x,y)=>x.name.localeCompare(y.name))
    .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");
}

/* -------------------- Routes -------------------- */
function onAddRoute() {
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
  const agentsById = new Map(state.agents.map(a => [a.id, a]));

  for (const r of state.routes.slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const tr = document.createElement("tr");

    tr.appendChild(tdInput(r.name, v => { r.name = v; saveRender(); }));

    const tdAgent = document.createElement("td");
    const sel = document.createElement("select");
    sel.className = "cell";
    sel.innerHTML = `<option value="">(fără)</option>` + state.agents
      .slice().sort((x,y)=>x.name.localeCompare(y.name))
      .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");
    sel.value = r.agentId || "";
    sel.addEventListener("change", () => {
      r.agentId = sel.value || null;
      saveRender();
    });
    tdAgent.appendChild(sel);
    tr.appendChild(tdAgent);

    const tdDel = document.createElement("td");
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
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

/* -------------------- Activities -------------------- */
function onAddActivity() {
  const name = $("#newActivityName").value.trim();
  if (!name) return alert("Completează activitatea.");
  state.activities.push({ id: uid(), name });
  $("#newActivityName").value = "";
  saveRender();
}

function renderActivities() {
  const tbody = $("#tblActivities tbody");
  tbody.innerHTML = "";
  for (const a of state.activities.slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const tr = document.createElement("tr");
    tr.appendChild(tdInput(a.name, v => { a.name = v; saveRender(); }));

    const tdDel = document.createElement("td");
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
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }

  renderActivitiesChips($("#visitActivities"), [], () => {});
}

/* -------------------- Planning selectors -------------------- */
function renderPlanningSelectors() {
  // clients
  const selC = $("#visitClient");
  selC.innerHTML = state.clients
    .slice().sort((a,b)=>a.name.localeCompare(b.name))
    .map(c => `<option value="${c.id}">${escapeHtml(c.name)}</option>`).join("");
  // agents
  const selA = $("#visitAgent");
  selA.innerHTML = `<option value="">(fără)</option>` + state.agents
    .slice().sort((a,b)=>a.name.localeCompare(b.name))
    .map(a => `<option value="${a.id}">${escapeHtml(a.name)}</option>`).join("");

  // routes
  const selR = $("#visitRoute");
  selR.innerHTML = `<option value="">(fără)</option>` + state.routes
    .slice().sort((a,b)=>a.name.localeCompare(b.name))
    .map(r => `<option value="${r.id}">${escapeHtml(r.name)}</option>`).join("");

  // default assign route->agent if route selected later handled on add
  renderActivitiesChips($("#visitActivities"), [], () => {});
}

function renderActivitiesChips(container, selectedIds, onChange) {
  container.innerHTML = "";
  const set = new Set(selectedIds || []);
  for (const a of state.activities.slice().sort((x,y)=>x.name.localeCompare(y.name))) {
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

/* -------------------- Visits (generate + manual) -------------------- */
function onGenerateVisits() {
  if (!state.clients.length) return alert("Nu ai clienți.");
  const { from, to } = getInterval();
  if (!from || !to) return alert("Alege intervalul (De la / Până la).");
  if (from > to) return alert("Interval invalid.");

  // generate for each month in interval based on client monthlyWeeks + monthlyCount
  const existingKey = new Set(state.visits.map(v => `${v.clientId}|${v.date}`)); // avoid duplicates same client+date
  let added = 0;

  const months = listMonthsInRange(from, to); // [{y,m}]
  for (const c of state.clients) {
    const weeks = (c.monthlyWeeks || []).slice().sort((a,b)=>a-b);
    const wantCount = clampInt(c.monthlyCount ?? weeks.length, 0, 50);

    for (const { y, m } of months) {
      const datesForWeeks = weeks
        .map(w => firstDateOfWeekOfMonth(y, m, w)) // Date
        .filter(d => d && d >= from && d <= to)
        .map(d => toISO(d));

      // keep only first N (monthlyCount)
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
  const agentId = $("#visitAgent").value || null;
  let routeId = $("#visitRoute").value || null;

  if (!date) return alert("Alege data.");
  if (!clientId) return alert("Alege client.");

  // chosen activities: read chips state by checking .chip.on
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

  // if route has agent and agent not set, auto-set
  if (routeId && !agentId) {
    const r = state.routes.find(x => x.id === routeId);
    if (r?.agentId) {
      // set in UI for clarity
      $("#visitAgent").value = r.agentId;
    }
  }

  state.visits.push({
    id: uid(),
    date,
    clientId,
    agentId: $("#visitAgent").value || null,
    routeId,
    activityIds: selectedActivityIds,
    otherActivity,
    details,
    obs
  });

  // clear optional fields
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

  const clientsById = new Map(state.clients.map(c => [c.id, c]));
  const agentsById  = new Map(state.agents.map(a => [a.id, a]));
  const routesById  = new Map(state.routes.map(r => [r.id, r]));
  const actsById    = new Map(state.activities.map(a => [a.id, a]));

  const visits = state.visits.filter(inInterval);

  for (const v of visits) {
    const tr = document.createElement("tr");

    const day = dayFromISO(v.date);
    tr.appendChild(tdText(day));

    tr.appendChild(tdDate(v.date, newDate => { v.date = newDate; saveRender(); }));

    // client select
    tr.appendChild(tdSelect(
      state.clients.slice().sort((a,b)=>a.name.localeCompare(b.name)),
      v.clientId,
      c => c.id,
      c => c.name,
      newId => { v.clientId = newId; saveRender(); }
    ));

    // agent select
    tr.appendChild(tdSelect(
      [{id:"", name:"(fără)"}, ...state.agents.slice().sort((a,b)=>a.name.localeCompare(b.name))],
      v.agentId || "",
      a => a.id,
      a => a.name,
      newId => { v.agentId = newId || null; saveRender(); }
    ));

    // route select
    tr.appendChild(tdSelect(
      [{id:"", name:"(fără)"}, ...state.routes.slice().sort((a,b)=>a.name.localeCompare(b.name))],
      v.routeId || "",
      r => r.id,
      r => r.name,
      newId => {
        v.routeId = newId || null;
        // if route has agent and visit agent empty, auto set
        const r = state.routes.find(x => x.id === v.routeId);
        if (r?.agentId && !v.agentId) v.agentId = r.agentId;
        saveRender();
      }
    ));

    // activities multi + other
    const tdAct = document.createElement("td");
    tdAct.appendChild(activityEditor(v, () => saveRender()));
    tr.appendChild(tdAct);

    // details
    tr.appendChild(tdInput(v.details || "", val => { v.details = val; saveRender(); }));

    // obs
    tr.appendChild(tdInput(v.obs || "", val => { v.obs = val; saveRender(); }));

    // delete
    const tdDel = document.createElement("td");
    const btn = document.createElement("button");
    btn.className = "btn danger";
    btn.textContent = "Șterge";
    btn.addEventListener("click", () => {
      if (!confirm("Ștergi vizita?")) return;
      state.visits = state.visits.filter(x => x.id !== v.id);
      saveRender();
    });
    tdDel.appendChild(btn);
    tr.appendChild(tdDel);

    tbody.appendChild(tr);
  }
}

function activityEditor(visit, onSave) {
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
  for (const a of state.activities.slice().sort((x,y)=>x.name.localeCompare(y.name))) {
    const chip = document.createElement("div");
    chip.className = "chip" + (set.has(a.id) ? " on" : "");
    chip.textContent = a.name;
    chip.addEventListener("click", () => {
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
  other.addEventListener("change", () => {
    visit.otherActivity = other.value.trim();
    onSave();
  });

  wrap.appendChild(box);
  wrap.appendChild(other);

  return wrap;
}

function onSortVisits() {
  const mode = $("#sortMode").value;
  const clientsById = new Map(state.clients.map(c => [c.id, c.name]));
  const agentsById  = new Map(state.agents.map(a => [a.id, a.name]));
  const routesById  = new Map(state.routes.map(r => [r.id, r.name]));

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

/* -------------------- Export Excel -------------------- */
function onExport() {
  const { from, to } = getInterval();
  const visits = state.visits.filter(v => {
    if (!from || !to) return true;
    const d = fromISO(v.date);
    return d >= from && d <= to;
  });

  const clientsById = new Map(state.clients.map(c => [c.id, c.name]));
  const actsById = new Map(state.activities.map(a => [a.id, a.name]));

  // Build rows: Zi | Data | Client | Activitate | Detalii | Observatii
  const rows = [
    ["Zi", "Data", "Client", "Activitate", "Detalii", "Observatii"]
  ];

  const sorted = visits.slice().sort((a,b)=> (a.date||"").localeCompare(b.date||""));

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

  // Optional: set column widths
  ws["!cols"] = [
    { wch: 10 }, // Zi
    { wch: 12 }, // Data
    { wch: 28 }, // Client
    { wch: 28 }, // Activitate
    { wch: 22 }, // Detalii
    { wch: 22 }, // Observatii
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

/* -------------------- Reports -------------------- */
function renderReport() {
  const { from, to } = getInterval();
  const visits = state.visits.filter(v => {
    if (!from || !to) return true;
    const d = fromISO(v.date);
    return d >= from && d <= to;
  });

  const uniqueClients = new Set(visits.map(v => v.clientId).filter(Boolean));

  const byAgent = new Map(); // agentId -> count
  const byRoute = new Map(); // routeId -> count
  for (const v of visits) {
    const a = v.agentId || "(fără)";
    const r = v.routeId || "(fără)";
    byAgent.set(a, (byAgent.get(a) || 0) + 1);
    byRoute.set(r, (byRoute.get(r) || 0) + 1);
  }

  const agentsById = new Map(state.agents.map(a => [a.id, a.name]));
  const routesById = new Map(state.routes.map(r => [r.id, r.name]));

  const stats = $("#reportStats");
  stats.innerHTML = "";

  stats.appendChild(statCard("Nr. vizite (interval)", String(visits.length)));
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

/* -------------------- Reset -------------------- */
function onResetAll() {
  if (!confirm("Reset total? Se șterg clienți, agenți, rute, activități și vizite.")) return;
  state = {
    clients: [],
    agents: [],
    routes: [],
    activities: [
      { id: uid(), name: "comanda" },
      { id: uid(), name: "incasare" },
      { id: uid(), name: "comanda + incasare" },
    ],
    visits: []
  };
  saveRender();
}

/* -------------------- Helpers UI -------------------- */
function tdInput(value, onChange) {
  const td = document.createElement("td");
  const input = document.createElement("input");
  input.className = "cell";
  input.value = value ?? "";
  input.addEventListener("change", () => onChange(input.value.trim()));
  td.appendChild(input);
  return td;
}

function tdText(value) {
  const td = document.createElement("td");
  td.textContent = value ?? "";
  return td;
}

function tdDate(iso, onChange) {
  const td = document.createElement("td");
  const input = document.createElement("input");
  input.type = "date";
  input.className = "cell";
  input.value = iso || "";
  input.addEventListener("change", () => onChange(input.value));
  td.appendChild(input);
  return td;
}

function tdSelect(items, selectedValue, getVal, getLabel, onChange) {
  const td = document.createElement("td");
  const sel = document.createElement("select");
  sel.className = "cell";
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

/* -------------------- Date helpers -------------------- */
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
  // iso: YYYY-MM-DD -> DD.MM.YYYY
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
  let m = from.getMonth(); // 0-based
  const endY = to.getFullYear();
  const endM = to.getMonth();
  while (y < endY || (y === endY && m <= endM)) {
    out.push({ y, m });
    m++;
    if (m > 11) { m = 0; y++; }
  }
  return out;
}

// Week-of-month: 1..5 based on day-of-month (1-7 => week1, 8-14=>week2 ...)
// pick first day of that week, then advance to Monday for consistency.
// If spills outside month, return null.
function firstDateOfWeekOfMonth(year, month0, weekNum) {
  const startDay = (weekNum - 1) * 7 + 1;
  const d = new Date(year, month0, startDay);
  if (d.getMonth() !== month0) return null;

  // move to Monday (1) within that week, but keep inside same week window
  // if startDay is e.g. 1st and it's Wednesday, go forward to Monday? that's backwards.
  // We'll choose the first day of the week block and then move forward to nearest weekday (Monday),
  // but if that goes past the 7-day block, keep original.
  const day = d.getDay(); // 0 Sun .. 6 Sat
  let deltaToMon = (1 - day + 7) % 7;
  const candidate = new Date(d);
  candidate.setDate(d.getDate() + deltaToMon);

  // ensure candidate is within same week block (startDay..startDay+6)
  if (candidate.getMonth() !== month0) return d;
  const candDay = candidate.getDate();
  if (candDay >= startDay && candDay <= startDay + 6) return candidate;
  return d;
}

/* -------------------- Persistence -------------------- */
function loadState() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return { clients: [], agents: [], routes: [], activities: [], visits: [] };
    const parsed = JSON.parse(raw);
    // minimal shape guard
    return {
      clients: Array.isArray(parsed.clients) ? parsed.clients : [],
      agents: Array.isArray(parsed.agents) ? parsed.agents : [],
      routes: Array.isArray(parsed.routes) ? parsed.routes : [],
      activities: Array.isArray(parsed.activities) ? parsed.activities : [],
      visits: Array.isArray(parsed.visits) ? parsed.visits : [],
    };
  } catch {
    return { clients: [], agents: [], routes: [], activities: [], visits: [] };
  }
}

function saveState(s) {
  localStorage.setItem(LS_KEY, JSON.stringify(s));
}

/* -------------------- Utils -------------------- */
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
