/**
 * ========================================================
 * LABCONTROL v2 - app.js
 * UI optimista: los datos aparecen al instante en pantalla
 * mientras se guardan en segundo plano en Google Sheets.
 * ========================================================
 */

const API_URL = "https://script.google.com/macros/s/AKfycbyQi0G8TGuLa_EBMpqVNEuCUl3lcTihRAJI1HlhR6BE2s_fwpSUBo_EZMtcu5Z7J9f2gA/exec";
const ADMIN_PASSWORD = "lab2025";
const POLL_INTERVAL_MS = 10000;

// Fallback si el Maestro no carga
const EXAMENES_DEFAULT = [
  "AC. FOLICO", "AC. URICO", "AFP", "ALBUMINA", "AMILASA", "BIL DIRECTA", "BIL TOTAL",
  "CA", "CA 125", "CA 19-9", "CEA", "CK", "CK-MB", "CL", "COL", "CORTISOL", "CRE",
  "DHEA-S", "ELECT. PROT", "ESTRADIOL", "FA", "FE", "FERRITINA", "FIBRINOGENO",
  "FSH", "GGT", "GLU", "GLOBULINA", "GOT", "GPT", "HBA1C", "HDL", "HEMOGRAMA",
  "INR", "INSULINA", "K", "LDH", "LDL", "LH", "LIPASA", "MG", "NA", "ORINA COMP",
  "P", "PCR", "PROGESTERONA", "PROLACTINA", "PROTEINAS TOT", "PROTROMBINA",
  "PSA", "PTH", "T3", "T4", "T4L", "TESTOSTERONA", "TG", "TP", "TROPO", "TSH",
  "TTPA", "URE", "VHS", "VIT B12", "VIT D"
];

const ESTADOS_CENTRO = [
  { value: "NO REVISADO :(", label: "NO REVISADO :(", css: "st-no-revisado" },
  { value: "REVISADO :)", label: "REVISADO :)", css: "st-revisado" },
  { value: "EN PROCESO", label: "EN PROCESO", css: "st-en-proceso" },
  { value: "TODO VALIDADO", label: "TODO VALIDADO ✓", css: "st-todo-validado" },
  { value: "NO LLEGÓ", label: "NO LLEGÓ", css: "st-no-llego" }
];

// ===================== ESTADO =====================
let datos = { errores: [], muestras: [], curvas: [], urgentes: [], recordatorios: [], custom: [], centros: [], maestro_examenes: [] };
let examenesAgregados = [];
let examenesEliminados = [];
let confirmCallback = null;
let pollTimer = null;
let pendingSaves = 0;

// ===================== INIT =====================
document.addEventListener("DOMContentLoaded", () => {
  autoDetectarDia();
  cargarDatos(true);
  iniciarPolling();
});

function autoDetectarDia() {
  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const hoy = dias[new Date().getDay()];
  const sel = document.getElementById("errDia");
  for (let i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === hoy) { sel.selectedIndex = i; break; }
  }
}

// ===================== TABS =====================
function switchTab(tabId) {
  document.querySelectorAll(".module-panel").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
  document.getElementById("panel-" + tabId).classList.add("active");
  document.querySelector(`[data-tab="${tabId}"]`).classList.add("active");
}

// ===================== API =====================
async function apiGet() {
  try {
    const r = await fetch(API_URL);
    return (await r.json()).data || null;
  } catch (e) { console.error("GET:", e); setSyncStatus("error"); return null; }
}

// POST en segundo plano - NO bloquea la UI
function apiPostBackground(payload) {
  pendingSaves++;
  showSaving();
  fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
    redirect: "follow"
  })
    .then(r => r.json())
    .then(j => {
      pendingSaves--;
      if (pendingSaves <= 0) { pendingSaves = 0; hideSaving(); }
      if (!j.success) showToast("Error al guardar: " + (j.error || ""), "error");
    })
    .catch(e => {
      pendingSaves--;
      if (pendingSaves <= 0) { pendingSaves = 0; hideSaving(); }
      showToast("Error de conexión", "error");
    });
}

// POST bloqueante (solo para cierre semanal)
async function apiPostBlocking(payload) {
  showLoading("Procesando...");
  try {
    const r = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });
    const j = await r.json();
    hideLoading();
    if (j.success) return j.data;
    throw new Error(j.error);
  } catch (e) { hideLoading(); showToast("Error: " + e.message, "error"); return null; }
}

// ===================== POLLING =====================
function iniciarPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(() => cargarDatos(false), POLL_INTERVAL_MS);
}

async function cargarDatos(showLoader) {
  if (showLoader) showLoading("Cargando datos...");
  const d = await apiGet();
  if (showLoader) hideLoading();
  if (d) {
    datos = d;
    poblarSelectsExamenes();
    renderErrores();
    renderMuestras();
    renderCurvas();
    renderUrgentes();
    renderRecordatorios();
    renderCustom();
    renderCentros();
    setSyncStatus("ok");
  }
}

function setSyncStatus(status) {
  const dot = document.getElementById("syncDot");
  const label = document.getElementById("syncLabel");
  if (status === "ok") { dot.style.background = "#22c55e"; dot.style.animation = "pulse 2s infinite"; label.textContent = "Sincronizado"; }
  else { dot.style.background = "#ef4444"; dot.style.animation = "none"; label.textContent = "Sin conexión"; }
}

// ===================== MAESTRO EXÁMENES =====================
function getExamenes() {
  return (datos.maestro_examenes && datos.maestro_examenes.length > 0) ? datos.maestro_examenes : EXAMENES_DEFAULT;
}

function poblarSelectsExamenes() {
  const lista = getExamenes();
  ["errExamenAgregar", "errExamenEliminar", "muestraExamen"].forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    const currentVal = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    lista.forEach(ex => { const o = document.createElement("option"); o.value = ex; o.textContent = ex; sel.appendChild(o); });
    sel.value = currentVal;
  });
}

// ===================== MODULE 1: ERRORES =====================
function addExamToList(tipo) {
  const selId = tipo === "agregar" ? "errExamenAgregar" : "errExamenEliminar";
  const arr = tipo === "agregar" ? examenesAgregados : examenesEliminados;
  const val = document.getElementById(selId).value;
  if (!val || arr.includes(val)) return;
  arr.push(val);
  document.getElementById(selId).value = "";
  renderExamPills(tipo === "agregar" ? "listaAgregados" : "listaEliminados", arr, tipo);
}

function removeExamFromList(tipo, exam) {
  const arr = tipo === "agregar" ? examenesAgregados : examenesEliminados;
  const idx = arr.indexOf(exam); if (idx > -1) arr.splice(idx, 1);
  renderExamPills(tipo === "agregar" ? "listaAgregados" : "listaEliminados", arr, tipo);
}

function renderExamPills(id, arr, tipo) {
  document.getElementById(id).innerHTML = arr.map(ex =>
    `<span class="exam-pill ${tipo === 'eliminar' ? 'exam-pill-red' : ''}">${ex}<button onclick="removeExamFromList('${tipo}','${ex}')" style="cursor:pointer;margin-left:2px;opacity:.7;">✕</button></span>`
  ).join("");
}

function submitError() {
  const dia = document.getElementById("errDia").value;
  const peticion = document.getElementById("errPeticion").value.trim();
  const usuario = document.getElementById("errUsuario").value.trim().toLowerCase();
  if (!dia || !peticion || !usuario) { showToast("Completa Día, Petición y Usuario", "error"); return; }
  if (!examenesAgregados.length && !examenesEliminados.length) { showToast("Selecciona al menos un examen", "error"); return; }

  const fecha = new Date().toLocaleDateString("es-CL");

  // UI Optimista: agregar a datos locales y renderizar AL INSTANTE
  examenesAgregados.forEach(ex => {
    datos.errores.push({ Fecha: fecha, "Día": dia, "Acción": "AGREGADO", "N° Petición": peticion, Examen: ex, Usuario: usuario });
    apiPostBackground({ action: "insert", sheet: "Errores", row: [fecha, dia, "AGREGADO", peticion, ex, usuario] });
  });
  examenesEliminados.forEach(ex => {
    datos.errores.push({ Fecha: fecha, "Día": dia, "Acción": "ELIMINADO", "N° Petición": peticion, Examen: ex, Usuario: usuario });
    apiPostBackground({ action: "insert", sheet: "Errores", row: [fecha, dia, "ELIMINADO", peticion, ex, usuario] });
  });

  renderErrores();
  document.getElementById("errPeticion").value = "";
  examenesAgregados = []; examenesEliminados = [];
  document.getElementById("listaAgregados").innerHTML = "";
  document.getElementById("listaEliminados").innerHTML = "";
  showToast("✓ Registro guardado", "success");
}

function renderErrores() {
  const container = document.getElementById("diasContainer");
  const dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];
  container.innerHTML = dias.map(dia => {
    const del = datos.errores.filter(e => (e["Día"] || e["D\u00eda"]) === dia);
    const ag = del.filter(e => (e["Acción"] || e["Acci\u00f3n"]) === "AGREGADO");
    const el = del.filter(e => (e["Acción"] || e["Acci\u00f3n"]) === "ELIMINADO");
    return `<div class="glass-card overflow-hidden">
      <div class="px-3 py-2 bg-slate-800/50 border-b border-slate-700">
        <h3 class="text-xs font-bold text-white">${dia}</h3>
        <span class="text-[10px] text-slate-500">${del.length} registro${del.length !== 1 ? 's' : ''}</span>
      </div>
      <div class="p-2.5 border-b border-slate-700/50">
        <div class="flex items-center gap-1 mb-1.5"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span><span class="text-[10px] font-bold text-emerald-400 uppercase">Agregados (${ag.length})</span></div>
        ${ag.length === 0 ? '<p class="text-[10px] text-slate-600 italic">—</p>' : ag.map(e => `<div class="flex items-center justify-between py-1 px-1.5 rounded mb-0.5 bg-emerald-900/10 border border-emerald-800/20"><div><span class="text-[10px] font-mono text-slate-300">#${e["N° Petición"] || e["N\u00b0 Petici\u00f3n"] || ''}</span> <span class="exam-pill ml-0.5">${e.Examen || ''}</span></div><span class="text-[9px] text-slate-500">${e.Usuario || ''}</span></div>`).join("")}
      </div>
      <div class="p-2.5">
        <div class="flex items-center gap-1 mb-1.5"><span class="w-1.5 h-1.5 rounded-full bg-red-500"></span><span class="text-[10px] font-bold text-red-400 uppercase">Eliminados (${el.length})</span></div>
        ${el.length === 0 ? '<p class="text-[10px] text-slate-600 italic">—</p>' : el.map(e => `<div class="flex items-center justify-between py-1 px-1.5 rounded mb-0.5 bg-red-900/10 border border-red-800/20"><div><span class="text-[10px] font-mono text-slate-300">#${e["N° Petición"] || e["N\u00b0 Petici\u00f3n"] || ''}</span> <span class="exam-pill exam-pill-red ml-0.5">${e.Examen || ''}</span></div><span class="text-[9px] text-slate-500">${e.Usuario || ''}</span></div>`).join("")}
      </div>
    </div>`;
  }).join("");
}

// ===================== MODULE 2: MUESTRAS =====================
function addMuestra() {
  const tipo = document.getElementById("muestraTipo").value;
  const pet = document.getElementById("muestraPeticion").value.trim();
  const exa = document.getElementById("muestraExamen").value;
  if (!pet || !exa) { showToast("Completa Petición y Examen", "error"); return; }
  const id = Date.now().toString();
  datos.muestras.push({ ID: id, Tipo: tipo, "N° Petición": pet, Examen: exa });
  renderMuestras();
  apiPostBackground({ action: "insert", sheet: "Pizarra_Muestras", row: [id, tipo, pet, exa] });
  document.getElementById("muestraPeticion").value = "";
  document.getElementById("muestraExamen").value = "";
  showToast("✓ Muestra registrada", "success");
}

function deleteMuestra(id) {
  datos.muestras = datos.muestras.filter(m => m.ID !== id);
  renderMuestras();
  apiPostBackground({ action: "delete", sheet: "Pizarra_Muestras", id });
}

function renderMuestras() {
  const r10 = datos.muestras.filter(m => m.Tipo && m.Tipo.includes("R10"));
  const c1 = datos.muestras.filter(m => m.Tipo && m.Tipo.includes("C1"));
  document.getElementById("tablaR10").innerHTML = r10.length === 0
    ? '<tr><td colspan="3" class="px-2 py-3 text-center text-[10px] text-slate-600 italic">Sin muestras</td></tr>'
    : r10.map(m => `<tr class="border-b border-slate-800 row-hover"><td class="px-2 py-1.5 font-mono">${m["N° Petición"] || ""}</td><td class="px-2 py-1.5"><span class="exam-pill">${m.Examen || ""}</span></td><td class="px-1"><button onclick="deleteMuestra('${m.ID}')" class="text-red-500 hover:text-red-400 text-[10px]">✕</button></td></tr>`).join("");
  document.getElementById("tablaC1").innerHTML = c1.length === 0
    ? '<tr><td colspan="3" class="px-2 py-3 text-center text-[10px] text-slate-600 italic">Sin muestras</td></tr>'
    : c1.map(m => `<tr class="border-b border-slate-800 row-hover"><td class="px-2 py-1.5 font-mono">${m["N° Petición"] || ""}</td><td class="px-2 py-1.5"><span class="exam-pill">${m.Examen || ""}</span></td><td class="px-1"><button onclick="deleteMuestra('${m.ID}')" class="text-red-500 hover:text-red-400 text-[10px]">✕</button></td></tr>`).join("");
}

// ===================== MODULE 2B: CURVAS =====================
function addCurva() {
  const pet = document.getElementById("curvaPeticion").value.trim();
  if (!pet) { showToast("Ingresa N° de Petición", "error"); return; }
  const id = Date.now().toString();
  datos.curvas.push({ ID: id, "N° Petición": pet, Validada: "No" });
  renderCurvas();
  apiPostBackground({ action: "insert", sheet: "Pizarra_Curvas", row: [id, pet, "No"] });
  document.getElementById("curvaPeticion").value = "";
  showToast("✓ Curva registrada", "success");
}

function toggleCurva(id, val) {
  const item = datos.curvas.find(c => c.ID === id);
  if (item) item.Validada = val;
  renderCurvas();
  apiPostBackground({ action: "update", sheet: "Pizarra_Curvas", id, updates: { Validada: val } });
}

function deleteCurva(id) {
  datos.curvas = datos.curvas.filter(c => c.ID !== id);
  renderCurvas();
  apiPostBackground({ action: "delete", sheet: "Pizarra_Curvas", id });
}

function renderCurvas() { renderValidationTable("tablaCurvas", datos.curvas, "Curva", "toggleCurva", "deleteCurva"); }

// ===================== MODULE 2C: URGENTES =====================
function addUrgente() {
  const pet = document.getElementById("urgentePeticion").value.trim();
  if (!pet) { showToast("Ingresa N° de Petición", "error"); return; }
  const id = Date.now().toString();
  datos.urgentes.push({ ID: id, "N° Petición": pet, Validada: "No" });
  renderUrgentes();
  apiPostBackground({ action: "insert", sheet: "Pizarra_Urgentes", row: [id, pet, "No"] });
  document.getElementById("urgentePeticion").value = "";
  showToast("✓ Urgente registrado", "success");
}

function toggleUrgente(id, val) {
  const item = datos.urgentes.find(c => c.ID === id);
  if (item) item.Validada = val;
  renderUrgentes();
  apiPostBackground({ action: "update", sheet: "Pizarra_Urgentes", id, updates: { Validada: val } });
}

function deleteUrgente(id) {
  datos.urgentes = datos.urgentes.filter(c => c.ID !== id);
  renderUrgentes();
  apiPostBackground({ action: "delete", sheet: "Pizarra_Urgentes", id });
}

function renderUrgentes() { renderValidationTable("tablaUrgentes", datos.urgentes, "Urgente", "toggleUrgente", "deleteUrgente"); }

// Shared renderer for Curvas and Urgentes tables
function renderValidationTable(tbodyId, items, label, toggleFn, deleteFn) {
  const tbody = document.getElementById(tbodyId);
  if (!items.length) {
    tbody.innerHTML = `<tr><td colspan="3" class="px-2 py-3 text-center text-[10px] text-slate-600 italic">Sin registros</td></tr>`;
    return;
  }
  tbody.innerHTML = items.map(c => {
    const v = c.Validada === "Sí" || c.Validada === "Si";
    const cls = v ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30";
    return `<tr class="border-b border-slate-800 row-hover">
      <td class="px-2 py-1.5 font-mono">${c["N° Petición"] || c["N\u00b0 Petici\u00f3n"] || ""}</td>
      <td class="px-2 py-1.5 text-center">
        <select onchange="${toggleFn}('${c.ID}',this.value)" class="text-[10px] font-bold px-1.5 py-0.5 rounded-full border ${cls}" style="background:transparent;cursor:pointer;">
          <option value="No" ${!v ? 'selected' : ''} style="background:#1e293b;color:#fca5a5;">❌ No</option>
          <option value="Sí" ${v ? 'selected' : ''} style="background:#1e293b;color:#6ee7b7;">✅ Sí</option>
        </select>
      </td>
      <td class="px-1"><button onclick="${deleteFn}('${c.ID}')" class="text-red-500 hover:text-red-400 text-[10px]">✕</button></td>
    </tr>`;
  }).join("");
}

// ===================== MODULE 2D: RECORDATORIOS =====================
function addRecordatorio() {
  const txt = document.getElementById("recordatorioTexto").value.trim();
  if (!txt) return;
  const id = Date.now().toString();
  const fecha = new Date().toLocaleString("es-CL");
  datos.recordatorios.push({ ID: id, Texto: txt, Usuario: "", Fecha: fecha });
  renderRecordatorios();
  apiPostBackground({ action: "insert", sheet: "Pizarra_Recordatorios", row: [id, txt, "", fecha] });
  document.getElementById("recordatorioTexto").value = "";
  showToast("✓ Nota añadida", "success");
}

function deleteRecordatorio(id) {
  datos.recordatorios = datos.recordatorios.filter(r => r.ID !== id);
  renderRecordatorios();
  apiPostBackground({ action: "delete", sheet: "Pizarra_Recordatorios", id });
}

function renderRecordatorios() {
  const container = document.getElementById("listaRecordatorios");
  if (!datos.recordatorios.length) { container.innerHTML = '<p class="text-[10px] text-slate-600 italic">Sin notas</p>'; return; }
  container.innerHTML = datos.recordatorios.map(r =>
    `<div class="flex items-start justify-between p-2 rounded-lg bg-slate-800/50 border border-slate-700/50">
      <div class="flex-1">
        <p class="text-xs text-slate-200">${escapeHtml(r.Texto)}</p>
        <p class="text-[9px] text-slate-500 mt-0.5">${r.Fecha || ""}</p>
      </div>
      <button onclick="deleteRecordatorio('${r.ID}')" class="text-red-500 hover:text-red-400 text-[10px] ml-2 mt-0.5 shrink-0">✕</button>
    </div>`
  ).join("");
}

// ===================== MODULE 2E: CUADRANTES CUSTOM =====================
function crearCuadrante() {
  const nombre = prompt("Nombre del nuevo cuadrante:");
  if (!nombre || !nombre.trim()) return;
  // Solo agregar un placeholder local - el cuadrante existe cuando tiene filas
  // Agregar una fila vacía para "crear" el cuadrante
  const id = Date.now().toString();
  datos.custom.push({ Cuadrante: nombre.trim(), ID: id, "N° Petición": "", Detalle: "(Cuadrante creado)" });
  renderCustom();
  apiPostBackground({ action: "insert", sheet: "Pizarra_Custom", row: [nombre.trim(), id, "", "(Cuadrante creado)"] });
  showToast("✓ Cuadrante '" + nombre.trim() + "' creado", "success");
}

function addCustomRow(cuadrante) {
  const petInput = document.getElementById("customPet_" + css(cuadrante));
  const detInput = document.getElementById("customDet_" + css(cuadrante));
  const pet = petInput ? petInput.value.trim() : "";
  const det = detInput ? detInput.value.trim() : "";
  if (!pet && !det) { showToast("Escribe algo", "error"); return; }
  const id = Date.now().toString();
  datos.custom.push({ Cuadrante: cuadrante, ID: id, "N° Petición": pet, Detalle: det });
  renderCustom();
  apiPostBackground({ action: "insert", sheet: "Pizarra_Custom", row: [cuadrante, id, pet, det] });
  showToast("✓ Añadido", "success");
}

function deleteCustomRow(id) {
  datos.custom = datos.custom.filter(c => c.ID !== id);
  renderCustom();
  apiPostBackground({ action: "delete", sheet: "Pizarra_Custom", id });
}

function deleteCustomCuadrante(cuadrante) {
  if (!confirm("¿Eliminar todo el cuadrante '" + cuadrante + "'?")) return;
  datos.custom = datos.custom.filter(c => c.Cuadrante !== cuadrante);
  renderCustom();
  apiPostBackground({ action: "delete_by_col", sheet: "Pizarra_Custom", column: "Cuadrante", value: cuadrante });
  showToast("Cuadrante eliminado", "success");
}

function css(str) { return str.replace(/[^a-zA-Z0-9]/g, "_"); }

function renderCustom() {
  const container = document.getElementById("customContainer");
  // Obtener nombres únicos de cuadrantes
  const cuadrantes = [...new Set(datos.custom.map(c => c.Cuadrante))];

  let html = cuadrantes.map(nombre => {
    const rows = datos.custom.filter(c => c.Cuadrante === nombre);
    const cid = css(nombre);
    return `<div class="glass-card p-4">
      <div class="flex items-center justify-between mb-2">
        <h3 class="text-xs font-bold text-white">${escapeHtml(nombre)}</h3>
        <button onclick="deleteCustomCuadrante('${escapeAttr(nombre)}')" class="text-red-500 hover:text-red-400 text-[10px]" title="Eliminar cuadrante">🗑</button>
      </div>
      <div class="flex gap-1 mb-2">
        <input type="text" id="customPet_${cid}" class="input-field text-xs" placeholder="Petición" style="width:35%;" />
        <input type="text" id="customDet_${cid}" class="input-field text-xs" placeholder="Detalle" style="width:50%;" />
        <button class="btn btn-xs btn-primary" onclick="addCustomRow('${escapeAttr(nombre)}')">+</button>
      </div>
      <div class="space-y-1">
        ${rows.filter(r => r["N° Petición"] || (r.Detalle && r.Detalle !== "(Cuadrante creado)")).map(r =>
      `<div class="flex items-center justify-between py-1 px-2 rounded bg-slate-800/50 border border-slate-700/30 text-[10px]">
            <div><span class="font-mono text-slate-300">${r["N° Petición"] || ""}</span>${r.Detalle && r.Detalle !== "(Cuadrante creado)" ? ' <span class="text-slate-400">— ' + escapeHtml(r.Detalle) + '</span>' : ''}</div>
            <button onclick="deleteCustomRow('${r.ID}')" class="text-red-500 hover:text-red-400">✕</button>
          </div>`
    ).join("")}
      </div>
    </div>`;
  }).join("");

  // Botón para crear nuevo
  html += `<div class="glass-card p-4 flex items-center justify-center border-2 border-dashed border-slate-600 cursor-pointer hover:border-indigo-500 transition-colors" onclick="crearCuadrante()">
    <p class="text-xs text-slate-500">+ Nuevo cuadrante</p>
  </div>`;

  container.innerHTML = html;
}

// ===================== MODULE 3: CENTROS =====================
function renderCentros() {
  const tbody = document.getElementById("tablaCentros");
  if (!datos.centros.length) { tbody.innerHTML = '<tr><td colspan="5" class="p-3 text-center text-xs text-slate-600">Sin centros</td></tr>'; return; }
  tbody.innerHTML = datos.centros.map(c => {
    const est = c.Estado || "NO REVISADO :(";
    const info = ESTADOS_CENTRO.find(e => e.value === est) || ESTADOS_CENTRO[0];
    return `<tr class="border-b border-slate-700/50 row-hover">
      <td class="px-3 py-2.5 font-semibold text-white text-xs">${c.Centro}</td>
      <td class="px-3 py-2.5 text-center">
        <select onchange="updateCentro('${escapeAttr(c.Centro)}','Estado',this.value)" class="status-badge ${info.css}" style="cursor:pointer;font-size:.65rem;padding:.2rem .5rem;border-radius:9999px;">
          ${ESTADOS_CENTRO.map(e => `<option value="${e.value}" ${e.value === est ? 'selected' : ''} style="background:#1e293b;color:#e2e8f0;">${e.label}</option>`).join("")}
        </select>
      </td>
      <td class="px-2 py-2.5 text-center"><input type="text" value="${c["Pdte Rev"] || ''}" placeholder="—" onchange="updateCentro('${escapeAttr(c.Centro)}','Pdte Rev',this.value)" class="input-field text-center text-[11px]" style="width:90px;" /></td>
      <td class="px-2 py-2.5 text-center"><input type="text" value="${c["Pdte Val"] || ''}" placeholder="—" onchange="updateCentro('${escapeAttr(c.Centro)}','Pdte Val',this.value)" class="input-field text-center text-[11px]" style="width:90px;" /></td>
      <td class="px-2 py-2.5 text-center"><input type="text" value="${c.Responsable || ''}" placeholder="—" maxlength="10" onchange="updateCentro('${escapeAttr(c.Centro)}','Responsable',this.value.toUpperCase())" class="input-field text-center text-[11px] font-bold" style="width:80px;text-transform:uppercase;" /></td>
    </tr>`;
  }).join("");
}

function updateCentro(centro, campo, valor) {
  const item = datos.centros.find(c => c.Centro === centro);
  if (item) item[campo] = valor;
  renderCentros();
  apiPostBackground({ action: "update", sheet: "Centros", id: centro, updates: { [campo]: valor } });
  showToast(`${centro}: actualizado`, "success");
}

// ===================== MODULE 4: ADMIN =====================
function cerrarSemana() {
  const pass = document.getElementById("adminPass").value;
  if (pass !== ADMIN_PASSWORD) { showToast("Contraseña incorrecta", "error"); return; }
  showConfirm("⚠️ Cerrar Semana", "Se exportarán todos los datos, se enviará correo y se limpiarán todas las tablas. ¿Continuar?", async () => {
    const r = await apiPostBlocking({ action: "cerrar_semana" });
    if (r) {
      showToast(typeof r === "string" ? r : "Semana cerrada", "success");
      document.getElementById("adminPass").value = "";
      await cargarDatos(false);
    }
  });
}

function descargarCSVLocal() {
  if (!datos.errores.length) { showToast("No hay datos para descargar", "info"); return; }
  const BOM = "\uFEFF";
  let csv = BOM + "Fecha,Día,Acción,N° Petición,Examen,Usuario\n";
  datos.errores.forEach(e => {
    const row = [e.Fecha || "", e["Día"] || e["D\u00eda"] || "", e["Acción"] || e["Acci\u00f3n"] || "", e["N° Petición"] || e["N\u00b0 Petici\u00f3n"] || "", e.Examen || "", e.Usuario || ""];
    csv += row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",") + "\n";
  });
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Errores_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  showToast("CSV descargado", "success");
}

// ===================== HELPERS =====================
function showToast(msg, type) {
  const t = document.getElementById("toast");
  t.className = `toast toast-${type} show`;
  t.textContent = msg;
  setTimeout(() => t.classList.remove("show"), 3000);
}

function showLoading(text) { document.getElementById("loadingText").textContent = text; document.getElementById("loadingOverlay").style.display = "flex"; }
function hideLoading() { document.getElementById("loadingOverlay").style.display = "none"; }

function showSaving() { document.getElementById("savingBadge").classList.add("show"); }
function hideSaving() { document.getElementById("savingBadge").classList.remove("show"); }

function showConfirm(title, msg, cb) {
  document.getElementById("confirmTitle").textContent = title;
  document.getElementById("confirmMsg").textContent = msg;
  confirmCallback = cb;
  document.getElementById("confirmModal").style.display = "flex";
}
function closeConfirm() { document.getElementById("confirmModal").style.display = "none"; confirmCallback = null; }
function executeConfirm() { if (confirmCallback) confirmCallback(); closeConfirm(); }

function escapeHtml(str) { const d = document.createElement("div"); d.textContent = str; return d.innerHTML; }
function escapeAttr(str) { return str.replace(/'/g, "\\'").replace(/"/g, "&quot;"); }
