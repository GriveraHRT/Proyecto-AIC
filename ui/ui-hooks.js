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

  /* @@INIT -------------------------------------------------- */
  var run = function () { inits.forEach(function (fn) { try { fn(); } catch (e) { console.error(e); } }); };
  if (doc.readyState === "loading") doc.addEventListener("DOMContentLoaded", run); else run();
})();
