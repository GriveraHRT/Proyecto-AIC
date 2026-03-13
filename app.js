/**
 * ========================================================
 * LABCONTROL v4 - app.js
 * Mejoras: llamas TODO VALIDADO, REV HASTA, reset centros,
 * stats separados, fechas pizarra, alertas, editar/eliminar errores
 * ========================================================
 */

const API_URL = "https://script.google.com/macros/s/AKfycbyQi0G8TGuLa_EBMpqVNEuCUl3lcTihRAJI1HlhR6BE2s_fwpSUBo_EZMtcu5Z7J9f2gA/exec";
const ADMIN_PASSWORD = "lab2025";
const POLL_INTERVAL_MS = 10000;
const ALERT_URGENTES_MS = 20 * 60 * 1000;   // 20 minutos
const ALERT_MUESTRAS_MS = 2 * 60 * 60 * 1000; // 2 horas

const EXAMENES_DEFAULT = ["AC. FOLICO", "AC. URICO", "AFP", "ALBUMINA", "AMILASA", "BIL DIRECTA", "BIL TOTAL", "CA", "CA 125", "CA 19-9", "CEA", "CK", "CK-MB", "CL", "COL", "CORTISOL", "CRE", "DHEA-S", "ELECT. PROT", "ESTRADIOL", "FA", "FE", "FERRITINA", "FIBRINOGENO", "FSH", "GGT", "GLU", "GLOBULINA", "GOT", "GPT", "HBA1C", "HDL", "HEMOGRAMA", "INR", "INSULINA", "K", "LDH", "LDL", "LH", "LIPASA", "MG", "NA", "ORINA COMP", "P", "PCR", "PROGESTERONA", "PROLACTINA", "PROTEINAS TOT", "PROTROMBINA", "PSA", "PTH", "T3", "T4", "T4L", "TESTOSTERONA", "TG", "TP", "TROPO", "TSH", "TTPA", "URE", "VHS", "VIT B12", "VIT D"];

const ESTADOS_CENTRO = [
  { value: "NO REVISADO :(", label: "NO REVISADO :(", css: "st-no-revisado" },
  { value: "REVISADO :)", label: "REVISADO :)", css: "st-revisado" },
  { value: "EN PROCESO", label: "EN PROCESO", css: "st-en-proceso" },
  { value: "TODO VALIDADO", label: "🔥 TODO VALIDADO", css: "st-todo-validado" },
  { value: "NO LLEGÓ", label: "NO LLEGÓ", css: "st-no-llego" }
];

let datos = { errores: [], muestras: [], curvas: [], urgentes: [], recordatorios: [], custom: [], centros: [], maestro_examenes: [], chat: [] };
let examenesAgregados = [], examenesEliminados = [];
let confirmCallback = null, pollTimer = null, pendingSaves = 0;
let alertTimerUrgentes = null, alertTimerMuestras = null;

// ===================== INIT =====================
document.addEventListener("DOMContentLoaded", () => {
  loadTheme();
  loadChatNick();
  autoDetectarDia();
  initFuzzy("fuzzyAgregar", "dropAgregar");
  initFuzzy("fuzzyEliminar", "dropEliminar");
  cargarDatos(true);
  iniciarPolling();
  iniciarAlertas();
});

// ===================== THEME =====================
function toggleTheme() {
  const html = document.documentElement;
  const next = html.getAttribute("data-theme") === "dark" ? "light" : "dark";
  html.setAttribute("data-theme", next);
  localStorage.setItem("labcontrol_theme", next);
  updateThemeLabel(next);
}
function loadTheme() {
  const saved = localStorage.getItem("labcontrol_theme");
  if (saved) document.documentElement.setAttribute("data-theme", saved);
  updateThemeLabel(saved || "dark");
}
function updateThemeLabel(theme) {
  const lbl = document.getElementById("themeLabel");
  if (lbl) lbl.textContent = theme === "dark" ? "🌙 Modo oscuro" : "☀️ Modo claro";
}

// ===================== TABS =====================
function switchTab(id) {
  document.querySelectorAll(".module-panel").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
  document.getElementById("panel-" + id).classList.add("active");
  document.querySelector(`[data-tab="${id}"]`).classList.add("active");
  if (id === "stats") renderStats();
  if (id === "chat") scrollChatBottom();
}

function autoDetectarDia() {
  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const sel = document.getElementById("errDia");
  const hoy = dias[new Date().getDay()];
  for (let i = 0; i < sel.options.length; i++) if (sel.options[i].value === hoy) { sel.selectedIndex = i; break; }
}

function fechaHoy() { return new Date().toLocaleDateString("es-CL"); }

// ===================== FUZZY SEARCH =====================
function initFuzzy(inputId, dropId) {
  const inp = document.getElementById(inputId);
  const drop = document.getElementById(dropId);
  inp.addEventListener("input", () => { renderFuzzyResults(inp, drop); });
  inp.addEventListener("focus", () => { renderFuzzyResults(inp, drop); });
  inp.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      const highlighted = drop.querySelector(".highlighted");
      if (highlighted) { inp.value = highlighted.dataset.val; drop.classList.remove("open"); }
      else { addExamFromFuzzy(inputId === "fuzzyAgregar" ? "agregar" : "eliminar"); }
    } else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault(); navigateFuzzy(drop, e.key === "ArrowDown" ? 1 : -1);
    } else if (e.key === "Escape") { drop.classList.remove("open"); }
  });
  document.addEventListener("click", (e) => { if (!inp.contains(e.target) && !drop.contains(e.target)) drop.classList.remove("open"); });
}

function renderFuzzyResults(inp, drop) {
  const query = inp.value.trim().toUpperCase();
  const lista = getExamenes();
  if (!query) {
    drop.innerHTML = lista.map(ex => `<div class="fuzzy-option" data-val="${ex}" onclick="selectFuzzy(this,'${inp.id}')">${ex}</div>`).join("");
    drop.classList.add("open");
    return;
  }
  const scored = lista.map(ex => ({ name: ex, score: fuzzyScore(query, ex) })).filter(x => x.score > 0).sort((a, b) => b.score - a.score);
  if (scored.length === 0) { drop.innerHTML = '<div class="fuzzy-option" style="opacity:.5;">Sin resultados</div>'; drop.classList.add("open"); return; }
  drop.innerHTML = scored.map((x, i) => `<div class="fuzzy-option ${i === 0 ? 'highlighted' : ''}" data-val="${x.name}" onclick="selectFuzzy(this,'${inp.id}')">${highlightMatch(x.name, query)}</div>`).join("");
  drop.classList.add("open");
}

function fuzzyScore(query, target) {
  const q = query.toUpperCase(), t = target.toUpperCase();
  if (t === q) return 100;
  if (t.startsWith(q)) return 90;
  if (t.includes(q)) return 80;
  let qi = 0, score = 0, consecutive = 0;
  for (let ti = 0; ti < t.length && qi < q.length; ti++) {
    if (t[ti] === q[qi]) { score += 10 + consecutive * 5; consecutive++; qi++; }
    else { consecutive = 0; }
  }
  if (qi < q.length) {
    const dist = levenshtein(q, t.substring(0, Math.min(t.length, q.length + 2)));
    if (dist <= 2) return 60 - dist * 10;
    return 0;
  }
  return score;
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) => {
    const row = new Array(n + 1).fill(0);
    row[0] = i; return row;
  });
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) {
    dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + (a[i - 1] !== b[j - 1] ? 1 : 0));
  }
  return dp[m][n];
}

function highlightMatch(text, query) {
  let result = "", qi = 0;
  for (let i = 0; i < text.length && qi < query.length; i++) {
    if (text[i].toUpperCase() === query[qi].toUpperCase()) { result += `<span class="fuzzy-match">${text[i]}</span>`; qi++; }
    else { result += text[i]; }
  }
  result += text.substring(result.replace(/<[^>]*>/g, '').length);
  return result;
}

function navigateFuzzy(drop, dir) {
  const opts = [...drop.querySelectorAll(".fuzzy-option[data-val]")];
  if (!opts.length) return;
  const cur = opts.findIndex(o => o.classList.contains("highlighted"));
  opts.forEach(o => o.classList.remove("highlighted"));
  const next = Math.max(0, Math.min(opts.length - 1, cur + dir));
  opts[next].classList.add("highlighted");
  opts[next].scrollIntoView({ block: "nearest" });
}

function selectFuzzy(el, inputId) {
  document.getElementById(inputId).value = el.dataset.val;
  el.closest(".fuzzy-dropdown").classList.remove("open");
}

function addExamFromFuzzy(tipo) {
  const inputId = tipo === "agregar" ? "fuzzyAgregar" : "fuzzyEliminar";
  const dropId = tipo === "agregar" ? "dropAgregar" : "dropEliminar";
  const inp = document.getElementById(inputId);
  const arr = tipo === "agregar" ? examenesAgregados : examenesEliminados;
  const val = inp.value.trim().toUpperCase();
  if (!val) return;
  const lista = getExamenes();
  const match = lista.find(e => e === val) || lista.find(e => e.startsWith(val));
  const final = match || val;
  if (arr.includes(final)) { showToast("Ya está en la lista", "info"); return; }
  arr.push(final);
  inp.value = "";
  document.getElementById(dropId).classList.remove("open");
  renderExamPills(tipo === "agregar" ? "listaAgregados" : "listaEliminados", arr, tipo);
}

// ===================== API =====================
async function apiGet() {
  try { const r = await fetch(API_URL); return (await r.json()).data || null; }
  catch (e) { console.error("GET:", e); setSyncStatus("error"); return null; }
}
function apiPostBg(payload) {
  pendingSaves++; showSaving();
  fetch(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload), redirect: "follow" })
    .then(r => r.json()).then(j => { pendingSaves--; if (pendingSaves <= 0) { pendingSaves = 0; hideSaving(); } if (!j.success) showToast("Error: " + (j.error || ""), "error"); })
    .catch(() => { pendingSaves--; if (pendingSaves <= 0) { pendingSaves = 0; hideSaving(); } showToast("Error de conexión", "error"); });
}
async function apiPostBlock(payload) {
  showLoading("Procesando...");
  try { const r = await fetch(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload), redirect: "follow" }); const j = await r.json(); hideLoading(); if (j.success) return j.data; throw new Error(j.error); }
  catch (e) { hideLoading(); showToast("Error: " + e.message, "error"); return null; }
}

// ===================== POLLING =====================
function iniciarPolling() { if (pollTimer) clearInterval(pollTimer); pollTimer = setInterval(() => cargarDatos(false), POLL_INTERVAL_MS); }
async function cargarDatos(loader) {
  if (loader) showLoading("Cargando...");
  const d = await apiGet();
  if (loader) hideLoading();
  if (d) { datos = d; poblarMuestraExamenes(); renderErrores(); renderMuestras(); renderCurvas(); renderUrgentes(); renderRecordatorios(); renderCustom(); renderCentros(); renderChat(); setSyncStatus("ok"); }
}
function setSyncStatus(s) { const d = document.getElementById("syncDot"), l = document.getElementById("syncLabel"); if (s === "ok") { d.style.background = "#22c55e"; d.style.animation = "pulse 2s infinite"; l.textContent = "Sincronizado"; } else { d.style.background = "#ef4444"; d.style.animation = "none"; l.textContent = "Sin conexión"; } }

function getExamenes() { return (datos.maestro_examenes && datos.maestro_examenes.length > 0) ? datos.maestro_examenes : EXAMENES_DEFAULT; }
function poblarMuestraExamenes() {
  const l = getExamenes();
  const sel = document.getElementById("muestraExamen"); if (!sel) return;
  const cv = sel.value; while (sel.options.length > 1) sel.remove(1);
  l.forEach(e => { const o = document.createElement("option"); o.value = e; o.textContent = e; sel.appendChild(o); }); sel.value = cv;
}

// ===================== ERRORES =====================
function removeExamFromList(t, ex) { const a = t === "agregar" ? examenesAgregados : examenesEliminados; const i = a.indexOf(ex); if (i > -1) a.splice(i, 1); renderExamPills(t === "agregar" ? "listaAgregados" : "listaEliminados", a, t); }
function renderExamPills(id, arr, t) { document.getElementById(id).innerHTML = arr.map(ex => `<span class="exam-pill ${t === 'eliminar' ? 'exam-pill-red' : ''}">${ex}<button onclick="removeExamFromList('${t}','${ex}')" style="cursor:pointer;margin-left:2px;opacity:.7;">✕</button></span>`).join(""); }

function submitError() {
  const dia = document.getElementById("errDia").value, pet = document.getElementById("errPeticion").value.trim(), usr = document.getElementById("errUsuario").value.trim().toLowerCase();
  if (!dia || !pet || !usr) { showToast("Completa Día, Petición y Usuario", "error"); return; }
  if (!examenesAgregados.length && !examenesEliminados.length) { showToast("Selecciona al menos un examen", "error"); return; }
  const fecha = fechaHoy();
  examenesAgregados.forEach(ex => { datos.errores.push({ Fecha: fecha, "Día": dia, "Acción": "AGREGADO", "N° Petición": pet, Examen: ex, Usuario: usr }); apiPostBg({ action: "insert", sheet: "Errores", row: [fecha, dia, "AGREGADO", pet, ex, usr] }); });
  examenesEliminados.forEach(ex => { datos.errores.push({ Fecha: fecha, "Día": dia, "Acción": "ELIMINADO", "N° Petición": pet, Examen: ex, Usuario: usr }); apiPostBg({ action: "insert", sheet: "Errores", row: [fecha, dia, "ELIMINADO", pet, ex, usr] }); });
  renderErrores(); document.getElementById("errPeticion").value = ""; examenesAgregados = []; examenesEliminados = []; document.getElementById("listaAgregados").innerHTML = ""; document.getElementById("listaEliminados").innerHTML = ""; showToast("✓ Guardado", "success");
}

function deleteError(idx) {
  const e = datos.errores[idx];
  if (!e) return;
  showConfirm("Eliminar registro", `¿Eliminar ${e["Acción"]} de ${e.Examen} (#${e["N° Petición"]})?`, () => {
    datos.errores.splice(idx, 1);
    renderErrores();
    // Delete from sheet: match by all columns
    apiPostBg({ action: "delete_by_col", sheet: "Errores", column: "N° Petición", value: e["N° Petición"] });
    // Re-insert remaining with same petition
    const remaining = datos.errores.filter(r => r["N° Petición"] === e["N° Petición"]);
    remaining.forEach(r => {
      apiPostBg({ action: "insert", sheet: "Errores", row: [r.Fecha, r["Día"], r["Acción"], r["N° Petición"], r.Examen, r.Usuario] });
    });
    showToast("Registro eliminado", "success");
  });
}

function editError(idx) {
  const e = datos.errores[idx];
  if (!e) return;
  const newPet = prompt("N° Petición:", e["N° Petición"]);
  if (newPet === null) return;
  const newExam = prompt("Examen:", e.Examen);
  if (newExam === null) return;
  const newUsr = prompt("Usuario:", e.Usuario);
  if (newUsr === null) return;
  // Update locally
  const oldPet = e["N° Petición"];
  e["N° Petición"] = newPet.trim() || oldPet;
  e.Examen = newExam.trim().toUpperCase() || e.Examen;
  e.Usuario = newUsr.trim().toLowerCase() || e.Usuario;
  renderErrores();
  // Delete old and re-insert all with old petition, plus updated one
  apiPostBg({ action: "delete_by_col", sheet: "Errores", column: "N° Petición", value: oldPet });
  const sameGroup = datos.errores.filter(r => r["N° Petición"] === oldPet || r === e);
  // Re-insert all that had the old petition
  datos.errores.filter(r => r["N° Petición"] === oldPet).forEach(r => {
    apiPostBg({ action: "insert", sheet: "Errores", row: [r.Fecha, r["Día"], r["Acción"], r["N° Petición"], r.Examen, r.Usuario] });
  });
  // If petition changed, also insert the new one
  if (e["N° Petición"] !== oldPet) {
    apiPostBg({ action: "insert", sheet: "Errores", row: [e.Fecha, e["Día"], e["Acción"], e["N° Petición"], e.Examen, e.Usuario] });
  }
  showToast("Registro actualizado", "success");
}

function renderErrores() {
  const c = document.getElementById("diasContainer"), dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];
  c.innerHTML = dias.map(dia => {
    const del = datos.errores.filter(e => (e["Día"] || e["D\u00eda"]) === dia);
    const ag = del.filter(e => (e["Acción"] || e["Acci\u00f3n"]) === "AGREGADO");
    const el = del.filter(e => (e["Acción"] || e["Acci\u00f3n"]) === "ELIMINADO");
    const findIdx = (item) => datos.errores.indexOf(item);
    return `<div class="glass-card overflow-hidden"><div class="px-2.5 py-2" style="border-bottom:1px solid var(--border);background:var(--bg-input);"><h3 class="text-xs font-bold">${dia}</h3><span class="text-[9px]" style="color:var(--text-dim);">${del.length} reg.</span></div>
    <div class="p-2" style="border-bottom:1px solid var(--border);"><div class="flex items-center gap-1 mb-1"><span class="w-1.5 h-1.5 rounded-full" style="background:#22c55e;"></span><span class="text-[9px] font-bold uppercase" style="color:var(--green-text);">Agregados (${ag.length})</span></div>
    ${ag.length === 0 ? '<p class="text-[9px] italic" style="color:var(--text-dim);">—</p>' : ag.map(e => `<div class="flex items-center justify-between py-0.5 px-1 rounded mb-0.5" style="background:rgba(16,185,129,.05);border:1px solid rgba(16,185,129,.15);"><div><span class="text-[9px] font-mono" style="color:var(--text);">#${e["N° Petición"] || ""}</span> <span class="exam-pill ml-0.5">${e.Examen || ""}</span></div><div class="flex items-center gap-1"><span class="text-[8px]" style="color:var(--text-dim);">${e.Usuario || ""}</span><button onclick="editError(${findIdx(e)})" class="text-[8px]" style="color:var(--accent);cursor:pointer;" title="Editar">✏️</button><button onclick="deleteError(${findIdx(e)})" class="text-[8px]" style="color:var(--red-text);cursor:pointer;" title="Eliminar">🗑</button></div></div>`).join("")}
    </div><div class="p-2"><div class="flex items-center gap-1 mb-1"><span class="w-1.5 h-1.5 rounded-full" style="background:#ef4444;"></span><span class="text-[9px] font-bold uppercase" style="color:var(--red-text);">Eliminados (${el.length})</span></div>
    ${el.length === 0 ? '<p class="text-[9px] italic" style="color:var(--text-dim);">—</p>' : el.map(e => `<div class="flex items-center justify-between py-0.5 px-1 rounded mb-0.5" style="background:rgba(239,68,68,.05);border:1px solid rgba(239,68,68,.15);"><div><span class="text-[9px] font-mono" style="color:var(--text);">#${e["N° Petición"] || ""}</span> <span class="exam-pill exam-pill-red ml-0.5">${e.Examen || ""}</span></div><div class="flex items-center gap-1"><span class="text-[8px]" style="color:var(--text-dim);">${e.Usuario || ""}</span><button onclick="editError(${findIdx(e)})" class="text-[8px]" style="color:var(--accent);cursor:pointer;" title="Editar">✏️</button><button onclick="deleteError(${findIdx(e)})" class="text-[8px]" style="color:var(--red-text);cursor:pointer;" title="Eliminar">🗑</button></div></div>`).join("")}
    </div></div>`;
  }).join("");
}

// ===================== MUESTRAS =====================
function addMuestra() {
  const t = document.getElementById("muestraTipo").value, p = document.getElementById("muestraPeticion").value.trim(), e = document.getElementById("muestraExamen").value;
  if (!p || !e) { showToast("Completa Petición y Examen", "error"); return; }
  const id = Date.now().toString(), fecha = fechaHoy();
  datos.muestras.push({ ID: id, Tipo: t, "N° Petición": p, Examen: e, Almacenada: "No", Fecha: fecha });
  renderMuestras();
  apiPostBg({ action: "insert", sheet: "Pizarra_Muestras", row: [id, t, p, e, "No", fecha] });
  document.getElementById("muestraPeticion").value = "";
  document.getElementById("muestraExamen").value = "";
  showToast("✓ Muestra", "success");
}
function deleteMuestra(id) { datos.muestras = datos.muestras.filter(m => m.ID !== id); renderMuestras(); apiPostBg({ action: "delete", sheet: "Pizarra_Muestras", id }); }
function toggleMuestraAlmacenada(id, val) {
  const item = datos.muestras.find(m => m.ID === id);
  if (item) item.Almacenada = val;
  renderMuestras();
  apiPostBg({ action: "update", sheet: "Pizarra_Muestras", id, updates: { Almacenada: val } });
}
function renderMuestras() {
  const sortByDate = (arr) => arr.sort((a, b) => {
    const da = parseFecha(a.Fecha), db = parseFecha(b.Fecha);
    return db - da;
  });
  const mk = (arr, tbId) => {
    const sorted = sortByDate([...arr]);
    document.getElementById(tbId).innerHTML = sorted.length === 0
      ? `<tr><td colspan="5" class="px-2 py-2.5 text-center text-[9px] italic" style="color:var(--text-dim);">Sin muestras</td></tr>`
      : sorted.map(m => {
        const alm = m.Almacenada === "Sí" || m.Almacenada === "Si";
        const cls = alm ? "st-revisado" : "st-no-revisado";
        return `<tr class="row-hover" style="border-bottom:1px solid var(--border);">
          <td class="px-2 py-1 text-[8px]" style="color:var(--text-dim);">${m.Fecha || ""}</td>
          <td class="px-2 py-1 font-mono text-xs">${m["N° Petición"] || ""}</td>
          <td class="px-2 py-1"><span class="exam-pill">${m.Examen || ""}</span></td>
          <td class="px-2 py-1 text-center">
            <select onchange="toggleMuestraAlmacenada('${m.ID}',this.value)" class="${cls}" style="font-size:.6rem;font-weight:700;padding:.15rem .4rem;border-radius:9999px;cursor:pointer;background:transparent;">
              <option value="No" ${!alm ? 'selected' : ''} style="background:var(--bg-input);color:var(--text);">❌ No</option>
              <option value="Sí" ${alm ? 'selected' : ''} style="background:var(--bg-input);color:var(--text);">✅ Sí</option>
            </select>
          </td>
          <td class="px-1"><button onclick="deleteMuestra('${m.ID}')" class="text-[9px]" style="color:var(--red-text);cursor:pointer;">✕</button></td>
        </tr>`;
      }).join("");
  };
  mk(datos.muestras.filter(m => m.Tipo && m.Tipo.includes("R10")), "tablaR10");
  mk(datos.muestras.filter(m => m.Tipo && m.Tipo.includes("C1")), "tablaC1");
}

// ===================== CURVAS / URGENTES =====================
function addCurva() {
  const p = document.getElementById("curvaPeticion").value.trim();
  if (!p) { showToast("Petición", "error"); return; }
  const id = Date.now().toString(), fecha = fechaHoy();
  datos.curvas.push({ ID: id, "N° Petición": p, Validada: "No", Fecha: fecha });
  renderCurvas();
  apiPostBg({ action: "insert", sheet: "Pizarra_Curvas", row: [id, p, "No", fecha] });
  document.getElementById("curvaPeticion").value = "";
  showToast("✓ Curva", "success");
}
function toggleCurva(id, v) { const i = datos.curvas.find(c => c.ID === id); if (i) i.Validada = v; renderCurvas(); apiPostBg({ action: "update", sheet: "Pizarra_Curvas", id, updates: { Validada: v } }); }
function deleteCurva(id) { datos.curvas = datos.curvas.filter(c => c.ID !== id); renderCurvas(); apiPostBg({ action: "delete", sheet: "Pizarra_Curvas", id }); }
function renderCurvas() { renderValTable("tablaCurvas", datos.curvas, "toggleCurva", "deleteCurva"); }

function addUrgente() {
  const p = document.getElementById("urgentePeticion").value.trim();
  if (!p) { showToast("Petición", "error"); return; }
  const id = Date.now().toString(), fecha = fechaHoy();
  datos.urgentes.push({ ID: id, "N° Petición": p, Validada: "No", Fecha: fecha });
  renderUrgentes();
  apiPostBg({ action: "insert", sheet: "Pizarra_Urgentes", row: [id, p, "No", fecha] });
  document.getElementById("urgentePeticion").value = "";
  showToast("✓ Urgente", "success");
}
function toggleUrgente(id, v) { const i = datos.urgentes.find(c => c.ID === id); if (i) i.Validada = v; renderUrgentes(); apiPostBg({ action: "update", sheet: "Pizarra_Urgentes", id, updates: { Validada: v } }); }
function deleteUrgente(id) { datos.urgentes = datos.urgentes.filter(c => c.ID !== id); renderUrgentes(); apiPostBg({ action: "delete", sheet: "Pizarra_Urgentes", id }); }
function renderUrgentes() { renderValTable("tablaUrgentes", datos.urgentes, "toggleUrgente", "deleteUrgente"); }

function renderValTable(tbId, items, toggleFn, delFn) {
  const tb = document.getElementById(tbId);
  const sorted = [...items].sort((a, b) => parseFecha(b.Fecha) - parseFecha(a.Fecha));
  if (!sorted.length) { tb.innerHTML = `<tr><td colspan="4" class="px-2 py-2.5 text-center text-[9px] italic" style="color:var(--text-dim);">Sin registros</td></tr>`; return; }
  tb.innerHTML = sorted.map(c => {
    const v = c.Validada === "Sí" || c.Validada === "Si"; const cls = v ? "st-revisado" : "st-no-revisado";
    return `<tr class="row-hover" style="border-bottom:1px solid var(--border);"><td class="px-2 py-1 text-[8px]" style="color:var(--text-dim);">${c.Fecha || ""}</td><td class="px-2 py-1 font-mono text-xs">${c["N° Petición"] || ""}</td><td class="px-2 py-1 text-center"><select onchange="${toggleFn}('${c.ID}',this.value)" class="${cls}" style="font-size:.6rem;font-weight:700;padding:.15rem .4rem;border-radius:9999px;cursor:pointer;background:transparent;"><option value="No" ${!v ? 'selected' : ''} style="background:var(--bg-input);color:var(--text);">❌ No</option><option value="Sí" ${v ? 'selected' : ''} style="background:var(--bg-input);color:var(--text);">✅ Sí</option></select></td><td class="px-1"><button onclick="${delFn}('${c.ID}')" class="text-[9px]" style="color:var(--red-text);cursor:pointer;">✕</button></td></tr>`;
  }).join("");
}

function parseFecha(f) {
  if (!f) return 0;
  // dd-mm-yyyy or dd/mm/yyyy
  const parts = f.split(/[-\/]/);
  if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
  return new Date(f).getTime() || 0;
}

// ===================== RECORDATORIOS =====================
function addRecordatorio() { const t = document.getElementById("recordatorioTexto").value.trim(); if (!t) return; const id = Date.now().toString(), f = new Date().toLocaleString("es-CL"); datos.recordatorios.push({ ID: id, Texto: t, Usuario: "", Fecha: f }); renderRecordatorios(); apiPostBg({ action: "insert", sheet: "Pizarra_Recordatorios", row: [id, t, "", f] }); document.getElementById("recordatorioTexto").value = ""; showToast("✓ Nota", "success"); }
function deleteRecordatorio(id) { datos.recordatorios = datos.recordatorios.filter(r => r.ID !== id); renderRecordatorios(); apiPostBg({ action: "delete", sheet: "Pizarra_Recordatorios", id }); }
function renderRecordatorios() { const c = document.getElementById("listaRecordatorios"); if (!datos.recordatorios.length) { c.innerHTML = '<p class="text-[9px] italic" style="color:var(--text-dim);">Sin notas</p>'; return; } c.innerHTML = datos.recordatorios.map(r => `<div class="flex items-start justify-between p-1.5 rounded-lg" style="background:var(--bg-input);border:1px solid var(--border);"><div class="flex-1"><p class="text-xs">${esc(r.Texto)}</p><p class="text-[8px]" style="color:var(--text-dim);">${r.Fecha || ""}</p></div><button onclick="deleteRecordatorio('${r.ID}')" class="text-[9px] ml-1.5 shrink-0" style="color:var(--red-text);cursor:pointer;">✕</button></div>`).join(""); }

// ===================== CUSTOM CUADRANTES =====================
function crearCuadrante() { const n = prompt("Nombre del cuadrante:"); if (!n || !n.trim()) return; const id = Date.now().toString(); datos.custom.push({ Cuadrante: n.trim(), ID: id, "N° Petición": "", Detalle: "(Creado)" }); renderCustom(); apiPostBg({ action: "insert", sheet: "Pizarra_Custom", row: [n.trim(), id, "", "(Creado)"] }); showToast("✓ Cuadrante creado", "success"); }
function addCustomRow(q) { const pi = document.getElementById("cp_" + css(q)), di = document.getElementById("cd_" + css(q)); const p = pi ? pi.value.trim() : "", d = di ? di.value.trim() : ""; if (!p && !d) { showToast("Escribe algo", "error"); return; } const id = Date.now().toString(); datos.custom.push({ Cuadrante: q, ID: id, "N° Petición": p, Detalle: d }); renderCustom(); apiPostBg({ action: "insert", sheet: "Pizarra_Custom", row: [q, id, p, d] }); showToast("✓", "success"); }
function deleteCustomRow(id) { datos.custom = datos.custom.filter(c => c.ID !== id); renderCustom(); apiPostBg({ action: "delete", sheet: "Pizarra_Custom", id }); }
function deleteCustomCuadrante(q) { if (!confirm("¿Eliminar '" + q + "'?")) return; datos.custom = datos.custom.filter(c => c.Cuadrante !== q); renderCustom(); apiPostBg({ action: "delete_by_col", sheet: "Pizarra_Custom", column: "Cuadrante", value: q }); showToast("Eliminado", "success"); }
function css(s) { return s.replace(/[^a-zA-Z0-9]/g, "_"); }
function renderCustom() {
  const c = document.getElementById("customContainer"); const qs = [...new Set(datos.custom.map(c => c.Cuadrante))];
  let h = qs.map(n => { const rows = datos.custom.filter(c => c.Cuadrante === n); const ci = css(n); return `<div class="glass-card p-3"><div class="flex items-center justify-between mb-1.5"><h3 class="text-xs font-bold">${esc(n)}</h3><button onclick="deleteCustomCuadrante('${escA(n)}')" class="text-[9px]" style="color:var(--red-text);cursor:pointer;">🗑</button></div><div class="flex gap-1 mb-1.5"><input type="text" id="cp_${ci}" class="input-field text-xs" placeholder="Petición" style="width:35%;" /><input type="text" id="cd_${ci}" class="input-field text-xs" placeholder="Detalle" style="width:50%;" /><button class="btn btn-xs btn-primary" onclick="addCustomRow('${escA(n)}')">+</button></div><div class="space-y-0.5">${rows.filter(r => r["N° Petición"] || (r.Detalle && r.Detalle !== "(Creado)")).map(r => `<div class="flex items-center justify-between py-0.5 px-1.5 rounded text-[9px]" style="background:var(--bg-input);border:1px solid var(--border);"><div><span class="font-mono">${r["N° Petición"] || ""}</span>${r.Detalle && r.Detalle !== "(Creado)" ? ' <span style="color:var(--text-muted);">— ' + esc(r.Detalle) + '</span>' : ''}</div><button onclick="deleteCustomRow('${r.ID}')" style="color:var(--red-text);cursor:pointer;">✕</button></div>`).join("")}</div></div>`; }).join("");
  h += `<div class="glass-card p-3 flex items-center justify-center border-2 border-dashed cursor-pointer hover:border-indigo-500 transition-colors" style="border-color:var(--border);" onclick="crearCuadrante()"><p class="text-xs" style="color:var(--text-dim);">+ Nuevo</p></div>`; c.innerHTML = h;
}

// ===================== CENTROS =====================
function renderCentros() {
  const tb = document.getElementById("tablaCentros"); if (!datos.centros.length) { tb.innerHTML = '<tr><td colspan="6" class="p-3 text-center text-xs" style="color:var(--text-dim);">Sin centros</td></tr>'; return; }
  tb.innerHTML = datos.centros.map(c => {
    const e = c.Estado || "NO REVISADO :("; const info = ESTADOS_CENTRO.find(s => s.value === e) || ESTADOS_CENTRO[0];
    return `<tr class="row-hover" style="border-bottom:1px solid var(--border);"><td class="px-2 py-2 font-semibold text-xs">${c.Centro}</td><td class="px-2 py-2 text-center"><select onchange="updateCentro('${escA(c.Centro)}','Estado',this.value)" class="${info.css}" style="cursor:pointer;font-size:.6rem;padding:.15rem .4rem;border-radius:9999px;position:relative;z-index:1;">${ESTADOS_CENTRO.map(s => `<option value="${s.value}" ${s.value === e ? 'selected' : ''} style="background:var(--bg-input);color:var(--text);">${s.label}</option>`).join("")}</select></td><td class="px-1 py-2 text-center"><input type="text" value="${c["Pdte Rev"] || ''}" placeholder="—" onchange="updateCentro('${escA(c.Centro)}','Pdte Rev',this.value)" class="input-field text-center text-[10px]" style="width:80px;" /></td><td class="px-1 py-2 text-center"><input type="text" value="${c["Pdte Val"] || ''}" placeholder="—" onchange="updateCentro('${escA(c.Centro)}','Pdte Val',this.value)" class="input-field text-center text-[10px]" style="width:80px;" /></td><td class="px-1 py-2 text-center"><input type="text" value="${c["Rev Hasta"] || ''}" placeholder="—" onchange="updateCentro('${escA(c.Centro)}','Rev Hasta',this.value)" class="input-field text-center text-[10px]" style="width:80px;" /></td><td class="px-1 py-2 text-center"><input type="text" value="${c.Responsable || ''}" placeholder="—" maxlength="10" onchange="updateCentro('${escA(c.Centro)}','Responsable',this.value.toUpperCase())" class="input-field text-center text-[10px] font-bold" style="width:70px;text-transform:uppercase;" /></td></tr>`;
  }).join("");
}
function updateCentro(c, f, v) { const i = datos.centros.find(x => x.Centro === c); if (i) i[f] = v; renderCentros(); apiPostBg({ action: "update", sheet: "Centros", id: c, updates: { [f]: v } }); showToast(`${c}: OK`, "success"); }

function resetCentros() {
  showConfirm("🔄 Reiniciar Centros", "Se borrarán Estado, Pdte Rev, Pdte Val, Rev Hasta y Responsable de TODOS los centros. ¿Continuar?", async () => {
    const r = await apiPostBlock({ action: "reset_centros" });
    if (r) {
      datos.centros.forEach(c => { c.Estado = "NO REVISADO :("; c["Pdte Rev"] = ""; c["Pdte Val"] = ""; c["Rev Hasta"] = ""; c.Responsable = ""; });
      renderCentros();
      showToast("Centros reiniciados", "success");
    }
  });
}

// ===================== CHAT =====================
function loadChatNick() { const n = localStorage.getItem("labcontrol_chat_nick"); if (n) { const el = document.getElementById("chatNickInput"); if (el) el.value = n; } }
function sendChat() {
  const nickEl = document.getElementById("chatNickInput"); const msgEl = document.getElementById("chatMsgInput");
  if (!nickEl || !msgEl) return;
  const nick = nickEl.value.trim(); const msg = msgEl.value.trim();
  if (!nick) { showToast("Escribe tu nombre", "error"); return; } if (!msg) return;
  localStorage.setItem("labcontrol_chat_nick", nick);
  const id = Date.now().toString(), fecha = new Date().toLocaleString("es-CL");
  datos.chat.push({ ID: id, Usuario: nick, Mensaje: msg, Fecha: fecha });
  renderChat(); apiPostBg({ action: "insert", sheet: "Chat", row: [id, nick, msg, fecha] });
  msgEl.value = ""; scrollChatBottom();
}
function renderChat() {
  const c = document.getElementById("chatMessages"); if (!c) return;
  const nickEl = document.getElementById("chatNickInput");
  const nick = nickEl ? (nickEl.value || "").trim().toLowerCase() : "";
  if (!datos.chat.length) { c.innerHTML = '<p class="text-center text-xs" style="color:var(--text-dim);padding:2rem;">Sin mensajes aún. ¡Sé el primero! 💬</p>'; return; }
  c.innerHTML = datos.chat.map(m => {
    const mine = m.Usuario.toLowerCase() === nick;
    return `<div style="display:flex;flex-direction:column;align-items:${mine ? 'flex-end' : 'flex-start'};margin-bottom:.5rem;">
    ${!mine ? `<div class="chat-sender">${esc(m.Usuario)}</div>` : ''}
    <div class="chat-bubble ${mine ? 'chat-mine' : 'chat-other'}">${esc(m.Mensaje)}</div>
    <div class="chat-time">${m.Fecha || ""}</div></div>`;
  }).join("");
}
function scrollChatBottom() { setTimeout(() => { const c = document.getElementById("chatMessages"); if (c) c.scrollTop = c.scrollHeight; }, 100); }

// ===================== ESTADÍSTICAS =====================
function renderStats() {
  const errs = datos.errores;
  const ag = errs.filter(e => (e["Acción"] || e["Acci\u00f3n"]) === "AGREGADO");
  const el = errs.filter(e => (e["Acción"] || e["Acci\u00f3n"]) === "ELIMINADO");
  document.getElementById("statTotal").textContent = errs.length;
  document.getElementById("statAgregados").textContent = ag.length;
  document.getElementById("statEliminados").textContent = el.length;

  // Top 15 usuarios
  const userCount = {}; errs.forEach(e => { const u = e.Usuario || "?"; userCount[u] = (userCount[u] || 0) + 1; });
  const topUsers = Object.entries(userCount).sort((a, b) => b[1] - a[1]).slice(0, 15);
  const maxU = topUsers.length ? topUsers[0][1] : 1;
  document.getElementById("statsTopUsuarios").innerHTML = topUsers.length === 0 ? '<p class="text-[10px] italic" style="color:var(--text-dim);">Sin datos</p>' : topUsers.map(([u, c]) => `<div class="flex items-center gap-2"><span class="text-[10px] font-mono font-bold" style="width:60px;color:var(--text);">${esc(u)}</span><div class="flex-1 rounded-full overflow-hidden" style="height:14px;background:var(--bg-input);"><div style="height:100%;width:${(c / maxU * 100).toFixed(0)}%;background:linear-gradient(90deg,#ef4444,#f97316);border-radius:9999px;"></div></div><span class="text-[10px] font-bold" style="color:var(--red-text);width:25px;text-align:right;">${c}</span></div>`).join("");

  // Top exámenes AGREGADOS
  const exAgCount = {}; ag.forEach(e => { const x = e.Examen || "?"; exAgCount[x] = (exAgCount[x] || 0) + 1; });
  const topExAg = Object.entries(exAgCount).sort((a, b) => b[1] - a[1]).slice(0, 15);
  const maxEAg = topExAg.length ? topExAg[0][1] : 1;
  document.getElementById("statsTopExAgregados").innerHTML = topExAg.length === 0 ? '<p class="text-[10px] italic" style="color:var(--text-dim);">Sin datos</p>' : topExAg.map(([x, c]) => `<div class="flex items-center gap-2"><span class="text-[10px] font-bold" style="width:80px;"><span class="exam-pill">${esc(x)}</span></span><div class="flex-1 rounded-full overflow-hidden" style="height:14px;background:var(--bg-input);"><div style="height:100%;width:${(c / maxEAg * 100).toFixed(0)}%;background:linear-gradient(90deg,#22c55e,#10b981);border-radius:9999px;"></div></div><span class="text-[10px] font-bold" style="color:var(--green-text);width:25px;text-align:right;">${c}</span></div>`).join("");

  // Top exámenes ELIMINADOS
  const exElCount = {}; el.forEach(e => { const x = e.Examen || "?"; exElCount[x] = (exElCount[x] || 0) + 1; });
  const topExEl = Object.entries(exElCount).sort((a, b) => b[1] - a[1]).slice(0, 15);
  const maxEEl = topExEl.length ? topExEl[0][1] : 1;
  document.getElementById("statsTopExEliminados").innerHTML = topExEl.length === 0 ? '<p class="text-[10px] italic" style="color:var(--text-dim);">Sin datos</p>' : topExEl.map(([x, c]) => `<div class="flex items-center gap-2"><span class="text-[10px] font-bold" style="width:80px;"><span class="exam-pill exam-pill-red">${esc(x)}</span></span><div class="flex-1 rounded-full overflow-hidden" style="height:14px;background:var(--bg-input);"><div style="height:100%;width:${(c / maxEEl * 100).toFixed(0)}%;background:linear-gradient(90deg,#ef4444,#dc2626);border-radius:9999px;"></div></div><span class="text-[10px] font-bold" style="color:var(--red-text);width:25px;text-align:right;">${c}</span></div>`).join("");
}

// ===================== ALERTAS =====================
function iniciarAlertas() {
  // Urgentes y Curvas: cada 20 min
  alertTimerUrgentes = setInterval(() => { checkAlertUrgentes(); checkAlertCurvas(); }, ALERT_URGENTES_MS);
  // Muestras no almacenadas: cada 2 horas
  alertTimerMuestras = setInterval(() => { checkAlertMuestras(); }, ALERT_MUESTRAS_MS);
  // También verificar al primer polling (5 seg después de carga)
  setTimeout(() => { checkAlertUrgentes(); checkAlertCurvas(); checkAlertMuestras(); }, 5000);
}

function checkAlertUrgentes() {
  const noVal = datos.urgentes.filter(u => u.Validada !== "Sí" && u.Validada !== "Si");
  if (noVal.length > 0) {
    const msgs = noVal.map(u => `🚨 Urgente #${u["N° Petición"]} NO validada`);
    showAlerts(msgs, "alert-item");
  }
}

function checkAlertCurvas() {
  const noVal = datos.curvas.filter(c => c.Validada !== "Sí" && c.Validada !== "Si");
  if (noVal.length > 0) {
    const msgs = noVal.map(c => `📈 Curva #${c["N° Petición"]} NO validada`);
    showAlerts(msgs, "alert-item");
  }
}

function checkAlertMuestras() {
  const noAlm = datos.muestras.filter(m => m.Almacenada !== "Sí" && m.Almacenada !== "Si");
  if (noAlm.length > 0) {
    const msgs = noAlm.map(m => `🧊 Muestra #${m["N° Petición"]} (${m.Examen}) NO almacenada`);
    showAlerts(msgs, "alert-item alert-item-warn");
  }
}

function showAlerts(msgs, cssClass) {
  const bar = document.getElementById("alertBar");
  bar.innerHTML = msgs.map(m => `<div class="${cssClass} alert-pulse">${m}</div>`).join("");
  setTimeout(() => { bar.innerHTML = ""; }, 15000);
}

// ===================== ADMIN =====================
function cerrarSemana() { const p = document.getElementById("adminPass").value; if (p !== ADMIN_PASSWORD) { showToast("Contraseña incorrecta", "error"); return; } showConfirm("⚠️ Cerrar Semana", "Exportará, enviará correo y limpiará todo. ¿Continuar?", async () => { const r = await apiPostBlock({ action: "cerrar_semana" }); if (r) { showToast(typeof r === "string" ? r : "Cerrado", "success"); document.getElementById("adminPass").value = ""; await cargarDatos(false); } }); }
function descargarCSVLocal() { if (!datos.errores.length) { showToast("Sin datos", "info"); return; } let csv = "\uFEFF" + "Fecha,Día,Acción,N° Petición,Examen,Usuario\n"; datos.errores.forEach(e => { csv += [e.Fecha || "", e["Día"] || "", e["Acción"] || "", e["N° Petición"] || "", e.Examen || "", e.Usuario || ""].map(v => `"${String(v).replace(/"/g, '""')}"`).join(",") + "\n"; }); const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8;" })); a.download = `Errores_${new Date().toISOString().slice(0, 10)}.csv`; a.click(); showToast("CSV descargado", "success"); }

// ===================== HELPERS =====================
function showToast(m, t) { const e = document.getElementById("toast"); e.className = `toast toast-${t} show`; e.textContent = m; setTimeout(() => e.classList.remove("show"), 3000); }
function showLoading(t) { document.getElementById("loadingText").textContent = t; document.getElementById("loadingOverlay").style.display = "flex"; }
function hideLoading() { document.getElementById("loadingOverlay").style.display = "none"; }
function showSaving() { document.getElementById("savingBadge").classList.add("show"); }
function hideSaving() { document.getElementById("savingBadge").classList.remove("show"); }
function showConfirm(t, m, cb) { document.getElementById("confirmTitle").textContent = t; document.getElementById("confirmMsg").textContent = m; confirmCallback = cb; document.getElementById("confirmModal").style.display = "flex"; }
function closeConfirm() { document.getElementById("confirmModal").style.display = "none"; confirmCallback = null; }
function executeConfirm() { if (confirmCallback) confirmCallback(); closeConfirm(); }
function esc(s) { const d = document.createElement("div"); d.textContent = s; return d.innerHTML; }
function escA(s) { return s.replace(/'/g, "\\'").replace(/"/g, "&quot;"); }
