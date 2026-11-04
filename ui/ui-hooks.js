/* =========================================================
   LabControl · ui-hooks.js — capa de presentación
   Se carga DESPUÉS de app.js. Envuelve funciones visuales:
   siempre llama a la original primero y solo añade clases,
   animaciones o sonido. No cambia datos ni sincronización.
   Organizado por momentos del storyboard (M1–M7); cada
   sección es independiente y se registra en onInit().
   ========================================================= */
(function () {
  "use strict";
  /* @@BASE ------------------------------------------------ */
  var doc = document, root = doc.documentElement;
  var $ = function (id) { return doc.getElementById(id); };
  var look = function () { return root.getAttribute("data-look"); };
  var reduced = function () { return window.matchMedia("(prefers-reduced-motion: reduce)").matches; };
  var inits = [];
  var onInit = function (fn) { inits.push(fn); };
  var wrap = function (name, fn) {
    var orig = window[name];
    if (typeof orig !== "function") return;
    window[name] = function () { return fn.call(this, orig, arguments); };
  };
  var restartClass = function (el, cls, ms) {
    if (!el) return;
    el.classList.remove(cls); void el.offsetWidth; el.classList.add(cls);
    if (ms) setTimeout(function () { el.classList.remove(cls); }, ms);
  };
  var ICON = function (n) { return '<svg class="ic"><use href="#i-' + n + '"/></svg>'; };
  var dataLoaded = false; // true tras la primera respuesta OK del servidor
  wrap("setSyncStatus", function (orig, a) { orig.apply(this, a); if (a[0] === "ok") dataLoaded = true; });

  /* Sonido sutil (solo look expresivo, apagado por defecto, interruptor visible) */
  var Sound = (function () {
    var ctx = null, on = false;
    try { on = localStorage.getItem("lc_sound") === "1"; } catch (e) { }
    function tone(freq, t0, dur, type) {
      var o = ctx.createOscillator(), g = ctx.createGain(), t = ctx.currentTime + t0;
      o.type = type || "sine"; o.frequency.setValueAtTime(freq, t);
      g.gain.setValueAtTime(0, t); g.gain.linearRampToValueAtTime(0.04, t + 0.008);
      g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
      o.connect(g).connect(ctx.destination); o.start(t); o.stop(t + dur + 0.02);
    }
    function play(kind) {
      if (!on || look() !== "expresivo") return;
      try {
        ctx = ctx || new (window.AudioContext || window.webkitAudioContext)();
        if (ctx.state === "suspended") ctx.resume();
        if (kind === "tick") tone(1320, 0, 0.05, "triangle");
        else if (kind === "ok") { tone(660, 0, 0.12); tone(990, 0.09, 0.16); }
        else if (kind === "error") tone(196, 0, 0.22, "triangle");
      } catch (e) { }
    }
    function sync() {
      var b = $("lcSoundToggle"); if (!b) return;
      b.setAttribute("aria-pressed", on ? "true" : "false");
      b.title = on ? "Sonidos activados" : "Sonidos apagados";
      b.setAttribute("aria-label", on ? "Desactivar sonidos" : "Activar sonidos");
    }
    function toggle() { on = !on; try { localStorage.setItem("lc_sound", on ? "1" : "0"); } catch (e) { } sync(); if (on) play("tick"); }
    return { play: play, sync: sync, toggle: toggle };
  })();
  onInit(function () {
    Sound.sync();
    var sb = $("lcSoundToggle"); if (sb) sb.addEventListener("click", Sound.toggle);
  });

  /* @@M1 · Carga inicial ------------------------------------ */
  onInit(function () {
    var intro = $("lcIntro"); if (!intro) return;
    if (root.classList.contains("no-intro") || reduced()) { intro.classList.add("is-done"); return; }
    var done = function () { intro.classList.add("is-done"); };
    intro.addEventListener("animationend", function (e) { if (e.target === intro) done(); });
    setTimeout(done, 1500); // tope duro (el omitir con clic/tecla vive en el <head>)
  });

  /* @@M2 · Cambio de pestaña -------------------------------- */
  var Tabs = (function () {
    var nav, ind;
    function place(animate) {
      nav = nav || $("lcTabs"); if (!nav) return;
      var act = nav.querySelector(".nav-tab.active"); if (!act) return;
      if (!ind) { ind = doc.createElement("span"); ind.className = "lc-tab-ind"; nav.prepend(ind); nav.classList.add("has-ind"); }
      ind.classList.toggle("no-anim", !animate);
      ind.style.width = act.offsetWidth + "px";
      ind.style.height = act.offsetHeight + "px";
      ind.style.transform = "translateX(" + act.offsetLeft + "px)";
      nav.classList.toggle("is-overflow", nav.scrollWidth > nav.clientWidth + 2);
      if (nav.scrollWidth > nav.clientWidth) {
        var target = act.offsetLeft - (nav.clientWidth - act.offsetWidth) / 2;
        nav.scrollTo({ left: target, behavior: animate && !reduced() ? "smooth" : "auto" });
      }
    }
    return { place: place };
  })();
  wrap("switchTab", function (orig, a) {
    var id = a[0], prev = doc.querySelector(".nav-tab.active");
    var r = orig.apply(this, a);
    var tab = doc.querySelector('.nav-tab[data-tab="' + id + '"]');
    Tabs.place(true);
    if (tab && tab !== prev) restartClass(tab, "is-entering", 400);
    try { history.replaceState(null, "", location.pathname + location.search + "#" + id); } catch (e) { }
    return r;
  });
  onInit(function () {
    Tabs.place(false);
    var rT; window.addEventListener("resize", function () { clearTimeout(rT); rT = setTimeout(function () { Tabs.place(false); }, 100); });
    if (doc.fonts && doc.fonts.ready) doc.fonts.ready.then(function () { Tabs.place(false); });
    // Abrir pestaña desde el hash (#pizarra, #stats…): útil para enlaces y pruebas
    var h = location.hash.replace("#", "");
    if (h && h !== "errores" && $("panel-" + h) && typeof switchTab === "function") switchTab(h);
  });

  /* @@M3 · Añadir / quitar examen --------------------------- */
  var pillName = function (el) { return el.firstChild ? el.firstChild.textContent.trim() : el.textContent.trim(); };
  wrap("renderExamPills", function (orig, a) {
    var cont = $(a[0]), before = {};
    if (cont) cont.querySelectorAll(".exam-pill").forEach(function (p) {
      before[pillName(p)] = { el: p, left: p.offsetLeft, top: p.offsetTop };
    });
    var r = orig.apply(this, a);
    if (!cont) return r;
    var now = {}, added = 0;
    cont.querySelectorAll(".exam-pill").forEach(function (p) {
      var n = pillName(p); now[n] = 1;
      if (!before[n]) { p.classList.add("is-new"); added++; }
    });
    if (!reduced()) Object.keys(before).forEach(function (n) {
      if (now[n]) return;
      // "fantasma": la lista real ya se actualizó; esto solo decora la salida
      var g = before[n].el.cloneNode(true);
      g.classList.add("lc-ghost"); g.classList.remove("is-new");
      g.style.left = before[n].left + "px"; g.style.top = before[n].top + "px";
      g.setAttribute("aria-hidden", "true");
      g.querySelectorAll("button").forEach(function (b) { b.remove(); });
      cont.appendChild(g);
      setTimeout(function () { g.remove(); }, 220);
    });
    if (added) { restartClass(cont.closest(".lc-exam-box"), "is-flash", 450); Sound.play("tick"); }
    return r;
  });
  var dupFlag = false;
  wrap("showToast", function (orig, a) { if (/Ya está en la lista/.test(a[0])) dupFlag = true; return orig.apply(this, a); });
  wrap("addExamFromFuzzy", function (orig, a) {
    var tipo = a[0], inp = $(tipo === "agregar" ? "fuzzyAgregar" : "fuzzyEliminar");
    var val = inp ? inp.value.trim().toUpperCase() : "";
    dupFlag = false;
    var r = orig.apply(this, a);
    if (val && dupFlag) {
      var lista = typeof getExamenes === "function" ? getExamenes() : [];
      var final = lista.find(function (e) { return e === val; }) || lista.find(function (e) { return e.indexOf(val) === 0; }) || val;
      var cont = $(tipo === "agregar" ? "listaAgregados" : "listaEliminados");
      if (cont) cont.querySelectorAll(".exam-pill").forEach(function (p) { if (pillName(p) === final) restartClass(p, "is-dup", 500); });
    }
    return r;
  });

  /* @@M4 · Guardar registro + toasts ------------------------ */
  wrap("showToast", function (orig, a) {
    orig.apply(this, a);
    var msg = String(a[0] || ""), type = a[1] || "info";
    var box = $("lcToasts"); if (!box) return;
    var el = doc.createElement("div");
    el.className = "lc-toast lc-toast--" + type;
    el.setAttribute("role", type === "error" ? "alert" : "status");
    el.innerHTML = ICON(type === "success" ? "check" : type === "error" ? "alert" : "info") + "<span></span>";
    el.querySelector("span").textContent = msg.replace(/^✓\s*/, "");
    var close = function () {
      if (el.classList.contains("is-out")) return;
      el.classList.add("is-out");
      setTimeout(function () { el.remove(); }, reduced() ? 0 : 160);
    };
    el.onclick = close;
    box.appendChild(el);
    while (box.children.length > 3) box.firstElementChild.remove();
    setTimeout(close, type === "error" ? 4500 : 3000);
    if (type === "error") Sound.play("error");
  });
  function flagInvalid(el) {
    if (!el) return;
    el.classList.add("lc-invalid");
    setTimeout(function () { el.classList.add("lc-invalid-fade"); el.classList.remove("lc-invalid"); }, 900);
    setTimeout(function () { el.classList.remove("lc-invalid-fade"); }, 1600);
  }
  wrap("submitError", function (orig, a) {
    // datos / examenes* son `let` globales de app.js (no cuelgan de window)
    var before = 0, noExams = false;
    try { before = datos.errores.length; noExams = !examenesAgregados.length && !examenesEliminados.length; } catch (e) { }
    var emptyDia = !$("errDia").value, emptyPet = !$("errPeticion").value.trim(), emptyUsr = !$("errUsuario").value.trim();
    var r = orig.apply(this, a);
    var after = before;
    try { after = datos.errores.length; } catch (e) { }
    var btn = $("lcSaveBtn");
    if (after > before) {
      restartClass(btn, "is-success", 750);
      Sound.play("ok");
      for (var i = before; i < after; i++) {
        var row = doc.querySelector('#diasContainer [data-idx="' + i + '"]');
        if (!row) continue;
        if (look() === "expresivo") row.style.animationDelay = ((i - before) * 40) + "ms, 0ms";
        row.classList.add("is-fresh");
      }
    } else {
      restartClass(btn, "is-shake", 320);
      if (emptyDia || emptyPet || emptyUsr) {
        if (emptyDia) flagInvalid($("errDia"));
        if (emptyPet) flagInvalid($("errPeticion"));
        if (emptyUsr) flagInvalid($("errUsuario"));
      } else if (noExams) {
        flagInvalid($("boxAgregados")); flagInvalid($("boxEliminados"));
      }
    }
    return r;
  });

  /* @@M5 · Chip de sincronización --------------------------- */
  var Status = (function () {
    var savedTimer = null, state = "ok";
    var LABEL = { ok: "Sincronizado", saving: "Guardando…", saved: "Guardado", error: "Sin conexión" };
    function set(s, label) {
      var chip = $("lcStatus"), text = $("lcStatusText"), logo = $("lcLogo");
      if (!chip) return;
      state = s;
      chip.setAttribute("data-state", s);
      text.textContent = label || LABEL[s];
      chip.title = label || LABEL[s];
      // El matraz pixel-art del logo acompaña el estado
      if (logo) logo.setAttribute("data-state", s === "saving" ? "working" : s === "saved" ? "done" : s === "error" ? "error" : "idle");
    }
    function beat() {
      if (look() !== "expresivo" || state !== "ok" || reduced()) return;
      restartClass($("lcStatus"), "is-beat", 700);
    }
    return {
      saving: function () { clearTimeout(savedTimer); set("saving"); },
      saved: function () {
        clearTimeout(savedTimer); set("saved");
        savedTimer = setTimeout(function () { if (state === "saved") set("ok"); }, 1500);
      },
      ok: function () { if (state === "saving" || state === "saved") return; var was = state; set("ok"); if (was === "ok") beat(); },
      error: function (label) { clearTimeout(savedTimer); set("error", label); }
    };
  })();
  wrap("showSaving", function (orig, a) { orig.apply(this, a); Status.saving(); });
  wrap("hideSaving", function (orig, a) { orig.apply(this, a); Status.saved(); });
  wrap("setSyncStatus", function (orig, a) { orig.apply(this, a); if (a[0] === "ok") Status.ok(); else Status.error("Sin conexión"); });
  wrap("showToast", function (orig, a) {
    orig.apply(this, a);
    if (a[1] === "error" && /conexi|^Error/i.test(String(a[0]))) Status.error("Error al guardar");
  });

  /* @@INIT -------------------------------------------------- */
  var run = function () { inits.forEach(function (fn) { try { fn(); } catch (e) { console.error(e); } }); };
  if (doc.readyState === "loading") doc.addEventListener("DOMContentLoaded", run); else run();
})();
