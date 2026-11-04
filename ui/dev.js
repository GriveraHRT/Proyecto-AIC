/* =========================================================
   LabControl · dev.js — SOLO PARA PRUEBAS LOCALES
   - En localhost/127.0.0.1/file:// simula la API de Google Sheets
     con datos ficticios (no escribe en la planilla real).
     Desactivar con ?demo=0. Simular red lenta: ?slow=1. Simular fallo: ?fail=1
   - Muestra un selector flotante de look (Sobrio / Expresivo).
   En producción (GitHub Pages) este archivo no hace nada.
   Se elimina al adoptar el look definitivo.
   ========================================================= */
(function () {
  var host = location.hostname;
  var isLocal = location.protocol === "file:" || host === "localhost" || host === "127.0.0.1";
  if (!isLocal) return;
  var p = new URLSearchParams(location.search);
  window.LC_DEV = true;

  // ---------------- Selector de look ----------------
  document.addEventListener("DOMContentLoaded", function () {
    var look = document.documentElement.getAttribute("data-look");
    var bar = document.createElement("div");
    bar.id = "lcDevBar";
    bar.innerHTML =
      '<span>Vista previa</span>' +
      '<button data-look="sobrio">A · Sobrio</button>' +
      '<button data-look="expresivo">B · Expresivo</button>' +
      (p.get("demo") === "0" ? '<em>API REAL</em>' : '<em>datos demo</em>');
    var css = document.createElement("style");
    css.textContent =
      "#lcDevBar{position:fixed;left:12px;bottom:12px;z-index:10001;display:flex;gap:4px;align-items:center;padding:4px;border-radius:999px;background:#111827;color:#e5e7eb;font:600 12px/1 system-ui,sans-serif;box-shadow:0 6px 20px rgba(0,0,0,.35)}" +
      "#lcDevBar span{padding:0 6px 0 8px;opacity:.7}#lcDevBar em{font-style:normal;padding:0 8px;color:#fbbf24}" +
      "#lcDevBar button{border:0;border-radius:999px;padding:7px 10px;background:transparent;color:inherit;cursor:pointer;font:inherit}" +
      "#lcDevBar button.on{background:#e5e7eb;color:#111827}";
    document.head.appendChild(css);
    bar.querySelectorAll("button").forEach(function (b) {
      if (b.dataset.look === look) b.classList.add("on");
      b.onclick = function () {
        var u = new URL(location.href);
        u.searchParams.set("look", b.dataset.look);
        u.searchParams.set("nointro", "");
        location.href = u.toString().replace("nointro=&", "nointro&").replace(/nointro=$/, "nointro");
      };
    });
    document.body.appendChild(bar);
  });

  if (p.get("demo") === "0") return;

  // ---------------- API simulada ----------------
  var hoy = new Date();
  var f = function (d) { var x = new Date(hoy); x.setDate(x.getDate() - d); return x.toLocaleDateString("es-CL"); };
  var diaHoy = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"][hoy.getDay()];
  var E = function (fe, dia, acc, pet, ex, us) { return { Fecha: fe, "Día": dia, "Acción": acc, "N° Petición": pet, Examen: ex, Usuario: us }; };

  var db = {
    errores: [
      E(f(3), "Lunes", "AGREGADO", "007187376", "TSH", "michellemt"),
      E(f(3), "Lunes", "ELIMINADO", "007187376", "T4L", "michellemt"),
      E(f(3), "Lunes", "AGREGADO", "007187411", "HEMOGRAMA", "jperez"),
      E(f(2), "Martes", "AGREGADO", "007188002", "PCR", "agonzalez"),
      E(f(2), "Martes", "ELIMINADO", "007188002", "VHS", "agonzalez"),
      E(f(2), "Martes", "ELIMINADO", "007188120", "FERRITINA", "jperez"),
      E(f(1), "Miércoles", "AGREGADO", "007189544", "HBA1C", "michellemt"),
      E(f(1), "Miércoles", "AGREGADO", "007189544", "GLU", "michellemt"),
      E(f(0), "Jueves", "ELIMINADO", "007190021", "CK-MB", "rsoto")
    ],
    muestras: [
      { ID: "m1", Tipo: "REFRIGERADA (R10)", "N° Petición": "007190110", Examen: "PTH", Almacenada: "No", Fecha: f(0) },
      { ID: "m2", Tipo: "REFRIGERADA (R10)", "N° Petición": "007189020", Examen: "VIT D", Almacenada: "Sí", Fecha: f(1) },
      { ID: "m3", Tipo: "CONGELADA (C1)", "N° Petición": "007188777", Examen: "INSULINA", Almacenada: "Sí", Fecha: f(2) }
    ],
    pendientes: [
      { ID: "p1", "N° Petición": "007190300", Examen: "TSH | T4L", Comentario: "Esperar segunda muestra", Procesada: "No", Fecha: f(0) },
      { ID: "p2", "N° Petición": "007190301", Examen: "PSA", Comentario: "", Procesada: "Sí", Fecha: f(0) }
    ],
    curvas: [{ ID: "c1", "N° Petición": "007190400", Validada: "No", Fecha: f(0) }],
    urgentes: [
      { ID: "u1", "N° Petición": "007190501", Validada: "Sí", Fecha: f(0) },
      { ID: "u2", "N° Petición": "007190502", Validada: "No", Fecha: f(0) }
    ],
    recordatorios: [{ ID: "r1", Texto: "Calibrar equipo de química a las 14:00", Usuario: "", Fecha: hoy.toLocaleString("es-CL") }],
    custom: [
      { Cuadrante: "Repetir mañana", ID: "q0", "N° Petición": "", Detalle: "(Creado)" },
      { Cuadrante: "Repetir mañana", ID: "q1", "N° Petición": "007190600", Detalle: "HEMÓLISIS" }
    ],
    centros: [
      { Centro: "CESFAM NORTE", Estado: "TODO VALIDADO", "Pdte Rev": "", "Pdte Val": "", "Rev Hasta": "", Responsable: "MT" },
      { Centro: "CESFAM SUR", Estado: "EN PROCESO", "Pdte Rev": "12", "Pdte Val": "4", "Rev Hasta": "0071899", Responsable: "JP" },
      { Centro: "HOSPITAL BASE", Estado: "REVISADO :)", "Pdte Rev": "", "Pdte Val": "2", "Rev Hasta": "", Responsable: "AG" },
      { Centro: "POSTA RURAL", Estado: "NO LLEGÓ", "Pdte Rev": "", "Pdte Val": "", "Rev Hasta": "", Responsable: "" },
      { Centro: "CECOSF ORIENTE", Estado: "NO REVISADO :(", "Pdte Rev": "", "Pdte Val": "", "Rev Hasta": "", Responsable: "" }
    ],
    maestro_examenes: [],
    chat: []
  };
  if (diaHoy !== "Jueves") db.errores[8]["Día"] = (["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"].indexOf(diaHoy) > -1 ? diaHoy : "Jueves");

  var SHEETS = {
    Errores: ["errores", ["Fecha", "Día", "Acción", "N° Petición", "Examen", "Usuario"]],
    Pizarra_Muestras: ["muestras", ["ID", "Tipo", "N° Petición", "Examen", "Almacenada", "Fecha"]],
    Pizarra_Pendientes: ["pendientes", ["ID", "N° Petición", "Examen", "Comentario", "Procesada", "Fecha"]],
    Pizarra_Curvas: ["curvas", ["ID", "N° Petición", "Validada", "Fecha"]],
    Pizarra_Urgentes: ["urgentes", ["ID", "N° Petición", "Validada", "Fecha"]],
    Pizarra_Recordatorios: ["recordatorios", ["ID", "Texto", "Usuario", "Fecha"]],
    Pizarra_Custom: ["custom", ["Cuadrante", "ID", "N° Petición", "Detalle"]],
    Centros: ["centros", ["Centro"]],
    Chat: ["chat", ["ID", "Usuario", "Mensaje", "Fecha"]]
  };

  function handle(body) {
    var s = SHEETS[body.sheet], key = s && s[0], list = key && db[key];
    switch (body.action) {
      case "insert": var o = {}; s[1].forEach(function (c, i) { o[c] = body.row[i]; }); list.push(o); return true;
      case "update": var it = list.find(function (r) { return r.ID === body.id || r.Centro === body.id; }); if (it) Object.assign(it, body.updates); return true;
      case "delete": db[key] = list.filter(function (r) { return r.ID !== body.id; }); return true;
      case "delete_by_col": db[key] = list.filter(function (r) { return String(r[body.column]) !== String(body.value); }); return true;
      case "clear_sheet": db[key] = []; return true;
      case "reset_centros": db.centros.forEach(function (c) { c.Estado = "NO REVISADO :("; c["Pdte Rev"] = c["Pdte Val"] = c["Rev Hasta"] = c.Responsable = ""; }); return true;
      case "get_historical_errors":
        var h = [], ex = ["TSH", "PCR", "GLU", "HEMOGRAMA", "CRE", "K", "NA", "VHS", "FERRITINA", "TROPO"], us = ["michellemt", "jperez", "agonzalez", "rsoto", "lmunoz"];
        var n = body.limit === -1 ? 120 : body.limit * 14;
        for (var i = 0; i < n; i++) h.push(E(f(7 + (i % 60)), "Lunes", i % 3 ? "AGREGADO" : "ELIMINADO", String(7100000 + i * 37), ex[(i * 7) % ex.length], us[(i * 3) % us.length]));
        return { errores: h };
      case "cerrar_semana": return "Semana cerrada (demo)";
      default: return true;
    }
  }

  var delay = p.has("slow") ? 3000 : 450;
  var realFetch = window.fetch.bind(window);
  window.fetch = function (url, opts) {
    if (typeof url !== "string" || url.indexOf("script.google.com") === -1) return realFetch(url, opts);
    return new Promise(function (resolve, reject) {
      setTimeout(function () {
        if (p.has("fail") && opts && opts.method === "POST") return reject(new TypeError("Demo: fallo simulado"));
        var payload;
        if (!opts || !opts.method || opts.method === "GET") payload = { success: true, data: JSON.parse(JSON.stringify(db)) };
        else payload = { success: true, data: handle(JSON.parse(opts.body)) };
        resolve(new Response(JSON.stringify(payload), { headers: { "Content-Type": "application/json" } }));
      }, (!opts || !opts.method) ? 350 : delay);
    });
  };
  console.info("[LabControl] Modo demo: la API está simulada. Usa ?demo=0 para la API real.");
})();
