/**
 * ========================================================
 * LABCONTROL - app.js
 * Lógica principal de la aplicación de Laboratorio Clínico
 * ========================================================
 */

// ===================== CONFIGURACIÓN =====================
const API_URL = "https://script.google.com/macros/s/AKfycbyQi0G8TGuLa_EBMpqVNEuCUl3lcTihRAJI1HlhR6BE2s_fwpSUBo_EZMtcu5Z7J9f2gA/exec";
const ADMIN_PASSWORD = "lab2025";   // Contraseña para cerrar semana
const POLL_INTERVAL_MS = 10000;     // Cada 10 segundos se actualiza

// Lista predefinida de exámenes (para evitar errores de escritura)
const EXAMENES = [
  "GLU", "LDL", "HDL", "COL", "TG", "TSH", "T3", "T4", "T4L",
  "HBA1C", "CRE", "URE", "AC. URICO", "BIL TOTAL", "BIL DIRECTA",
  "GOT", "GPT", "GGT", "FA", "LDH", "CK", "CK-MB", "TROPO",
  "NA", "K", "CL", "CA", "MG", "P", "FE", "FERRITINA", "VIT B12",
  "AC. FOLICO", "VIT D", "PTH", "PCR", "VHS", "HEMOGRAMA",
  "PROTROMBINA", "TTPA", "FIBRINOGENO", "INR", "ORINA COMP",
  "PSA", "CEA", "AFP", "CA 19-9", "CA 125", "INSULINA",
  "CORTISOL", "PROLACTINA", "FSH", "LH", "ESTRADIOL",
  "PROGESTERONA", "TESTOSTERONA", "DHEA-S", "AMILASA", "LIPASA",
  "PROTEINAS TOT", "ALBUMINA", "GLOBULINA", "TP", "ELECT. PROT"
].sort();

// Estados posibles para los centros
const ESTADOS_CENTRO = [
  { value: "NO REVISADO :(", label: "NO REVISADO :(", css: "st-no-revisado" },
  { value: "REVISADO :)", label: "REVISADO :)", css: "st-revisado" },
  { value: "EN PROCESO", label: "EN PROCESO", css: "st-en-proceso" },
  { value: "TODO VALIDADO", label: "TODO VALIDADO ✓", css: "st-todo-validado" },
  { value: "NO LLEGÓ", label: "NO LLEGÓ", css: "st-no-llego" }
];

// ===================== ESTADO LOCAL =====================
let datosLocales = { errores: [], muestras: [], curvas: [], centros: [] };
let examenesAgregados = [];
let examenesEliminados = [];
let confirmCallback = null;
let pollTimer = null;

// ===================== INICIALIZACIÓN =====================
document.addEventListener("DOMContentLoaded", () => {
  cargarUsuarioGuardado();
  poblarSelectsExamenes();
  autoDetectarDia();
  cargarDatos(true);
  iniciarPolling();
});

function cargarUsuarioGuardado() {
  const saved = localStorage.getItem("labcontrol_user");
  if (saved) document.getElementById("globalUser").value = saved;
  document.getElementById("globalUser").addEventListener("change", (e) => {
    localStorage.setItem("labcontrol_user", e.target.value.toUpperCase());
    e.target.value = e.target.value.toUpperCase();
  });
}

function poblarSelectsExamenes() {
  const selects = ["errExamenAgregar", "errExamenEliminar", "muestraExamen"];
  selects.forEach(id => {
    const sel = document.getElementById(id);
    // Limpiar opciones menos la primera
    while (sel.options.length > 1) sel.remove(1);
    EXAMENES.forEach(ex => {
      const opt = document.createElement("option");
      opt.value = ex;
      opt.textContent = ex;
      sel.appendChild(opt);
    });
  });
}

function autoDetectarDia() {
  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const hoy = dias[new Date().getDay()];
  const sel = document.getElementById("errDia");
  for (let i = 0; i < sel.options.length; i++) {
    if (sel.options[i].value === hoy) { sel.selectedIndex = i; break; }
  }
}

// ===================== TAB NAVIGATION =====================
function switchTab(tabId) {
  document.querySelectorAll(".module-panel").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
  document.getElementById("panel-" + tabId).classList.add("active");
  document.querySelector(`[data-tab="${tabId}"]`).classList.add("active");
}

// ===================== COMUNICACIÓN CON API =====================
async function apiGet() {
  try {
    const res = await fetch(API_URL);
    const json = await res.json();
    if (json.success) return json.data;
    throw new Error(json.error || "Error desconocido");
  } catch (err) {
    console.error("Error GET:", err);
    setSyncStatus("error");
    return null;
  }
}

async function apiPost(payload) {
  try {
    showLoading("Guardando...");
    const res = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
      redirect: "follow"
    });
    const json = await res.json();
    hideLoading();
    if (json.success) return json.data;
    throw new Error(json.error || "Error desconocido");
  } catch (err) {
    hideLoading();
    console.error("Error POST:", err);
    showToast("Error al guardar: " + err.message, "error");
    return null;
  }
}

// ===================== POLLING (AUTO-REFRESH) =====================
function iniciarPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(() => cargarDatos(false), POLL_INTERVAL_MS);
}

async function cargarDatos(showLoader) {
  if (showLoader) showLoading("Cargando datos del laboratorio...");
  const data = await apiGet();
  if (showLoader) hideLoading();

  if (data) {
    datosLocales = data;
    renderErrores();
    renderMuestras();
    renderCurvas();
    renderCentros();
    setSyncStatus("ok");
  }
}

function setSyncStatus(status) {
  const dot = document.getElementById("syncDot");
  const label = document.getElementById("syncLabel");
  if (status === "ok") {
    dot.style.background = "#22c55e";
    label.textContent = "Sincronizado";
  } else if (status === "error") {
    dot.style.background = "#ef4444";
    dot.style.animation = "none";
    label.textContent = "Sin conexión";
  }
}

// ===================== MODULE 1: ERRORES =====================

function addExamToList(tipo) {
  const selectId = tipo === "agregar" ? "errExamenAgregar" : "errExamenEliminar";
  const listId = tipo === "agregar" ? "listaAgregados" : "listaEliminados";
  const arr = tipo === "agregar" ? examenesAgregados : examenesEliminados;

  const sel = document.getElementById(selectId);
  const val = sel.value;
  if (!val) return;
  if (arr.includes(val)) { showToast("Ese examen ya está en la lista", "info"); return; }

  arr.push(val);
  sel.value = "";
  renderExamPills(listId, arr, tipo);
}

function removeExamFromList(tipo, exam) {
  const listId = tipo === "agregar" ? "listaAgregados" : "listaEliminados";
  const arr = tipo === "agregar" ? examenesAgregados : examenesEliminados;
  const idx = arr.indexOf(exam);
  if (idx > -1) arr.splice(idx, 1);
  renderExamPills(listId, arr, tipo);
}

function renderExamPills(containerId, arr, tipo) {
  const container = document.getElementById(containerId);
  container.innerHTML = arr.map(ex =>
    `<span class="exam-pill ${tipo === 'eliminar' ? 'exam-pill-red' : ''}">${ex}
      <button onclick="removeExamFromList('${tipo}','${ex}')" style="cursor:pointer;margin-left:2px;opacity:0.7;">✕</button>
    </span>`
  ).join("");
}

async function submitError() {
  const dia = document.getElementById("errDia").value;
  const peticion = document.getElementById("errPeticion").value.trim();
  const usuario = (document.getElementById("errUsuario").value || document.getElementById("globalUser").value).toUpperCase().trim();

  if (!dia || !peticion || !usuario) {
    showToast("Completa todos los campos obligatorios (Día, N° Petición y Usuario)", "error");
    return;
  }
  if (examenesAgregados.length === 0 && examenesEliminados.length === 0) {
    showToast("Debes seleccionar al menos un examen agregado o eliminado", "error");
    return;
  }

  const fecha = new Date().toLocaleDateString("es-CL");
  const promises = [];

  // Crear una fila por cada examen agregado
  examenesAgregados.forEach(ex => {
    promises.push(apiPost({
      action: "insert",
      sheet: "Errores",
      row: [fecha, dia, "AGREGADO", peticion, ex, usuario]
    }));
  });

  // Crear una fila por cada examen eliminado
  examenesEliminados.forEach(ex => {
    promises.push(apiPost({
      action: "insert",
      sheet: "Errores",
      row: [fecha, dia, "ELIMINADO", peticion, ex, usuario]
    }));
  });

  showLoading("Guardando registros...");
  await Promise.all(promises);
  hideLoading();

  // Limpiar formulario
  document.getElementById("errPeticion").value = "";
  examenesAgregados = [];
  examenesEliminados = [];
  document.getElementById("listaAgregados").innerHTML = "";
  document.getElementById("listaEliminados").innerHTML = "";

  showToast("¡Registro guardado correctamente!", "success");
  await cargarDatos(false);
}

function renderErrores() {
  const container = document.getElementById("diasContainer");
  const dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

  container.innerHTML = dias.map(dia => {
    const erroresDia = datosLocales.errores.filter(e => e["Día"] === dia || e["D\u00eda"] === dia);
    const agregados = erroresDia.filter(e => e["Acción"] === "AGREGADO" || e["Acci\u00f3n"] === "AGREGADO");
    const eliminados = erroresDia.filter(e => e["Acción"] === "ELIMINADO" || e["Acci\u00f3n"] === "ELIMINADO");

    return `
      <div class="day-column flex-shrink-0">
        <div class="glass-card overflow-hidden">
          <div class="px-4 py-2.5 bg-slate-800/50 border-b border-slate-700">
            <h3 class="text-sm font-bold text-white">${dia}</h3>
            <span class="text-xs text-slate-500">${erroresDia.length} registro${erroresDia.length !== 1 ? 's' : ''}</span>
          </div>

          <!-- Agregados -->
          <div class="p-3 border-b border-slate-700/50">
            <div class="flex items-center gap-1.5 mb-2">
              <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
              <span class="text-xs font-bold text-emerald-400 uppercase tracking-wider">Agregados (${agregados.length})</span>
            </div>
            ${agregados.length === 0 ? '<p class="text-xs text-slate-600 italic">Sin registros</p>' :
              agregados.map(e => `
                <div class="flex items-center justify-between py-1.5 px-2 rounded mb-1 bg-emerald-900/10 border border-emerald-800/20 row-hover">
                  <div>
                    <span class="text-xs font-mono text-slate-300">#${e["N° Petición"] || e["N\u00b0 Petici\u00f3n"] || ''}</span>
                    <span class="exam-pill ml-1">${e["Examen"] || ''}</span>
                  </div>
                  <span class="text-[10px] text-slate-500">${e["Usuario"] || ''}</span>
                </div>
              `).join("")}
          </div>

          <!-- Eliminados -->
          <div class="p-3">
            <div class="flex items-center gap-1.5 mb-2">
              <span class="w-2 h-2 rounded-full bg-red-500"></span>
              <span class="text-xs font-bold text-red-400 uppercase tracking-wider">Eliminados (${eliminados.length})</span>
            </div>
            ${eliminados.length === 0 ? '<p class="text-xs text-slate-600 italic">Sin registros</p>' :
              eliminados.map(e => `
                <div class="flex items-center justify-between py-1.5 px-2 rounded mb-1 bg-red-900/10 border border-red-800/20 row-hover">
                  <div>
                    <span class="text-xs font-mono text-slate-300">#${e["N° Petición"] || e["N\u00b0 Petici\u00f3n"] || ''}</span>
                    <span class="exam-pill exam-pill-red ml-1">${e["Examen"] || ''}</span>
                  </div>
                  <span class="text-[10px] text-slate-500">${e["Usuario"] || ''}</span>
                </div>
              `).join("")}
          </div>
        </div>
      </div>`;
  }).join("");
}

// ===================== MODULE 2: PIZARRA =====================

async function addMuestra() {
  const tipo = document.getElementById("muestraTipo").value;
  const peticion = document.getElementById("muestraPeticion").value.trim();
  const examen = document.getElementById("muestraExamen").value;

  if (!peticion || !examen) {
    showToast("Completa N° Petición y Examen", "error");
    return;
  }

  const id = Date.now().toString();
  const result = await apiPost({
    action: "insert",
    sheet: "Pizarra_Muestras",
    row: [id, tipo, peticion, examen]
  });

  if (result) {
    document.getElementById("muestraPeticion").value = "";
    document.getElementById("muestraExamen").value = "";
    showToast("Muestra registrada", "success");
    await cargarDatos(false);
  }
}

async function deleteMuestra(id) {
  const result = await apiPost({ action: "delete", sheet: "Pizarra_Muestras", id: id });
  if (result) {
    showToast("Muestra eliminada", "success");
    await cargarDatos(false);
  }
}

function renderMuestras() {
  const r10 = datosLocales.muestras.filter(m => m["Tipo"] && m["Tipo"].includes("R10"));
  const c1 = datosLocales.muestras.filter(m => m["Tipo"] && m["Tipo"].includes("C1"));

  document.getElementById("tablaR10").innerHTML = r10.length === 0
    ? '<tr><td colspan="3" class="px-3 py-4 text-center text-xs text-slate-600 italic">Sin muestras R10</td></tr>'
    : r10.map(m => `
      <tr class="border-b border-slate-800 row-hover">
        <td class="px-3 py-2 font-mono text-sm">${m["N° Petición"] || m["N\u00b0 Petici\u00f3n"] || ''}</td>
        <td class="px-3 py-2"><span class="exam-pill">${m["Examen"] || ''}</span></td>
        <td class="px-1 py-2"><button onclick="deleteMuestra('${m['ID']}')" class="text-red-500 hover:text-red-400 text-xs p-1" title="Eliminar">✕</button></td>
      </tr>`
    ).join("");

  document.getElementById("tablaC1").innerHTML = c1.length === 0
    ? '<tr><td colspan="3" class="px-3 py-4 text-center text-xs text-slate-600 italic">Sin muestras C1</td></tr>'
    : c1.map(m => `
      <tr class="border-b border-slate-800 row-hover">
        <td class="px-3 py-2 font-mono text-sm">${m["N° Petición"] || m["N\u00b0 Petici\u00f3n"] || ''}</td>
        <td class="px-3 py-2"><span class="exam-pill">${m["Examen"] || ''}</span></td>
        <td class="px-1 py-2"><button onclick="deleteMuestra('${m['ID']}')" class="text-red-500 hover:text-red-400 text-xs p-1" title="Eliminar">✕</button></td>
      </tr>`
    ).join("");
}

// ===================== MODULE 2B: CURVAS =====================

async function addCurva() {
  const peticion = document.getElementById("curvaPeticion").value.trim();
  if (!peticion) { showToast("Ingresa un N° de Petición", "error"); return; }

  const id = Date.now().toString();
  const result = await apiPost({
    action: "insert",
    sheet: "Pizarra_Curvas",
    row: [id, peticion, "No"]
  });

  if (result) {
    document.getElementById("curvaPeticion").value = "";
    showToast("Curva registrada", "success");
    await cargarDatos(false);
  }
}

async function toggleCurvaValidada(id, nuevoValor) {
  await apiPost({
    action: "update",
    sheet: "Pizarra_Curvas",
    id: id,
    updates: { "Validada": nuevoValor }
  });
  await cargarDatos(false);
}

async function deleteCurva(id) {
  const result = await apiPost({ action: "delete", sheet: "Pizarra_Curvas", id: id });
  if (result) {
    showToast("Curva eliminada", "success");
    await cargarDatos(false);
  }
}

function renderCurvas() {
  const tbody = document.getElementById("tablaCurvas");
  if (datosLocales.curvas.length === 0) {
    tbody.innerHTML = '<tr><td colspan="3" class="px-3 py-4 text-center text-xs text-slate-600 italic">Sin curvas registradas</td></tr>';
    return;
  }

  tbody.innerHTML = datosLocales.curvas.map(c => {
    const validada = c["Validada"] === "Sí" || c["Validada"] === "Si" || c["Validada"] === "S\u00ed";
    const statusClass = validada ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30";
    return `
      <tr class="border-b border-slate-800 row-hover">
        <td class="px-3 py-2 font-mono text-sm">${c["N° Petición"] || c["N\u00b0 Petici\u00f3n"] || ''}</td>
        <td class="px-3 py-2 text-center">
          <select onchange="toggleCurvaValidada('${c['ID']}', this.value)"
            class="text-xs font-bold px-2 py-1 rounded-full border ${statusClass}" style="background:transparent;cursor:pointer;">
            <option value="No" ${!validada ? 'selected' : ''} style="background:#1e293b;color:#fca5a5;">❌ No</option>
            <option value="Sí" ${validada ? 'selected' : ''} style="background:#1e293b;color:#6ee7b7;">✅ Sí</option>
          </select>
        </td>
        <td class="px-1 py-2"><button onclick="deleteCurva('${c['ID']}')" class="text-red-500 hover:text-red-400 text-xs p-1" title="Eliminar">✕</button></td>
      </tr>`;
  }).join("");
}

// ===================== MODULE 3: CENTROS =====================

function renderCentros() {
  const tbody = document.getElementById("tablaCentros");
  if (datosLocales.centros.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="px-3 py-4 text-center text-xs text-slate-600 italic">Sin centros configurados</td></tr>';
    return;
  }

  tbody.innerHTML = datosLocales.centros.map((c, idx) => {
    const estadoActual = c["Estado"] || "NO REVISADO :(";
    const estadoInfo = ESTADOS_CENTRO.find(e => e.value === estadoActual) || ESTADOS_CENTRO[0];

    return `
      <tr class="border-b border-slate-700/50 row-hover">
        <td class="px-3 py-3">
          <span class="font-semibold text-white">${c["Centro"]}</span>
        </td>
        <td class="px-3 py-3 text-center">
          <select onchange="updateCentro('${c['Centro']}', 'Estado', this.value)"
            class="status-badge ${estadoInfo.css}" style="cursor:pointer;font-size:0.7rem;">
            ${ESTADOS_CENTRO.map(e => `<option value="${e.value}" ${e.value === estadoActual ? 'selected' : ''} style="background:#1e293b;color:#e2e8f0;">${e.label}</option>`).join("")}
          </select>
        </td>
        <td class="px-3 py-3 text-center">
          <input type="text" value="${c["Pdte Rev"] || ''}" placeholder="—"
            onchange="updateCentro('${c['Centro']}', 'Pdte Rev', this.value)"
            class="input-field text-center text-xs" style="width:100px;" />
        </td>
        <td class="px-3 py-3 text-center">
          <input type="text" value="${c["Pdte Val"] || ''}" placeholder="—"
            onchange="updateCentro('${c['Centro']}', 'Pdte Val', this.value)"
            class="input-field text-center text-xs" style="width:100px;" />
        </td>
        <td class="px-3 py-3 text-center">
          <input type="text" value="${c["Responsable"] || ''}" placeholder="—" maxlength="5"
            onchange="updateCentro('${c['Centro']}', 'Responsable', this.value.toUpperCase())"
            class="input-field text-center text-xs font-bold" style="width:70px;text-transform:uppercase;" />
        </td>
      </tr>`;
  }).join("");
}

async function updateCentro(centro, campo, valor) {
  await apiPost({
    action: "update",
    sheet: "Centros",
    id: centro,
    updates: { [campo]: valor }
  });
  showToast(`${centro}: ${campo} actualizado`, "success");
  await cargarDatos(false);
}

// ===================== MODULE 4: ADMIN =====================

function cerrarSemana() {
  const pass = document.getElementById("adminPass").value;
  if (pass !== ADMIN_PASSWORD) {
    showToast("Contraseña incorrecta", "error");
    return;
  }
  showConfirm(
    "⚠️ Cerrar Semana",
    "Esta acción exportará todos los errores de la semana, enviará un correo a grivera@hospitaldetalca.cl, y limpiará la tabla. ¿Estás seguro?",
    async () => {
      const result = await apiPost({ action: "cerrar_semana" });
      if (result) {
        showToast(typeof result === "string" ? result : "Semana cerrada exitosamente", "success");
        document.getElementById("adminPass").value = "";
        await cargarDatos(false);
      }
    }
  );
}

function descargarCSVLocal() {
  if (datosLocales.errores.length === 0) {
    showToast("No hay datos de errores para descargar", "info");
    return;
  }

  const BOM = "\uFEFF";
  let csv = BOM + "Fecha,Día,Acción,N° Petición,Examen,Usuario\n";
  datosLocales.errores.forEach(e => {
    const row = [
      e["Fecha"] || "",
      e["Día"] || e["D\u00eda"] || "",
      e["Acción"] || e["Acci\u00f3n"] || "",
      e["N° Petición"] || e["N\u00b0 Petici\u00f3n"] || "",
      e["Examen"] || "",
      e["Usuario"] || ""
    ].map(v => `"${String(v).replace(/"/g, '""')}"`);
    csv += row.join(",") + "\n";
  });

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Errores_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast("Archivo CSV descargado", "success");
}

// ===================== UI HELPERS =====================

function showToast(message, type = "info") {
  const toast = document.getElementById("toast");
  toast.className = `toast toast-${type} show`;
  toast.textContent = message;
  setTimeout(() => toast.classList.remove("show"), 3500);
}

function showLoading(text) {
  document.getElementById("loadingText").textContent = text || "Cargando...";
  document.getElementById("loadingOverlay").style.display = "flex";
}

function hideLoading() {
  document.getElementById("loadingOverlay").style.display = "none";
}

function showConfirm(title, msg, callback) {
  document.getElementById("confirmTitle").textContent = title;
  document.getElementById("confirmMsg").textContent = msg;
  confirmCallback = callback;
  document.getElementById("confirmModal").style.display = "flex";
}

function closeConfirm() {
  document.getElementById("confirmModal").style.display = "none";
  confirmCallback = null;
}

function executeConfirm() {
  if (confirmCallback) confirmCallback();
  closeConfirm();
}
