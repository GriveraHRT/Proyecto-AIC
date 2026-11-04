# LabControl — Design Storyboard (rediseño visual)

> Estado: **borrador para aprobación**. No se ha tocado código.
> Alcance: solo presentación. `app.js`, IDs, nombres de funciones, estructura de datos y sincronización con Google Sheets **no se modifican**.

---

## 0. Cómo se entregan las dos opciones

| | **A · "Clínico"** (sobria) | **B · "Señal"** (expresiva) |
|---|---|---|
| Idea | Una hoja clínica bien impresa: calma, orden, cero ruido. | Un monitor de paciente: el sistema "late" y responde con pulso y color. |
| Motor de animación | Solo CSS + un archivo JS de ganchos visuales (0 KB de librerías) | CSS + **GSAP 3 por CDN** (~28 KB gz) para la intro y las estadísticas |
| Intro | 0,6 s | 1,2 s (máx. 1,5 s) |
| Sonido | No | Opcional, apagado por defecto |
| Glass / blur | Eliminado | Solo en el header |

**Flujo local:**
1. Rama `rediseno-visual` (master y GitHub Pages quedan intactos).
2. Las dos opciones conviven en la rama y se eligen con `?look=sobrio` / `?look=expresivo`. Comparten la capa de ganchos y difieren en CSS y tokens de movimiento.
3. Se sirven en local con `python -m http.server 5500` y se abren dos pestañas lado a lado.
4. Cuando elijas, un commit borra la opción descartada y el selector, y se integra a `master`.

**Archivos nuevos (no se edita `app.js`):**
- `ui/tokens.css` (base común), `ui/look-sobrio.css`, `ui/look-expresivo.css`
- `ui/ui-hooks.js`: se carga **después** de `app.js`. Envuelve funciones visuales existentes (`showToast`, `showSaving`, `hideSaving`, `setSyncStatus`, `switchTab`, `toggleTheme`, `renderExamPills`, `renderStats`, `submitError`). Llama siempre a la original y solo añade clases o animaciones; la lógica no cambia.
- `ui/icons.svg`: sprite de iconos propio.
- `index.html`: solo markup visual (iconos, chip de estado, contenedor de intro, interruptor de sonido, enlaces a CSS/JS).

---

## 1. Diagnóstico de la página actual

1. **Estructura:** página única (`index.html`) con unas 650 líneas de CSS inline + Tailwind Play CDN. Tiene 6 paneles `section.module-panel` que se alternan con `switchTab()`, y `app.js` genera casi todo el contenido como *template strings* con estilos inline.
2. **Ciclo de datos:** `cargarDatos()` hace polling cada 10 s y **re-renderiza con `innerHTML`** errores, pizarra, centros y estadísticas. Cualquier animación de entrada ingenua se repetiría cada 10 s (ver principio P1).
3. **Tipografía diminuta:** hay muchísimo texto en 8–10 px (`text-[8px]`, `text-[9px]`). Es ilegible en turnos largos o en monitores de baja densidad.
4. **Colores fuera del sistema:** muchos colores van fijos en el JS y el HTML (`#22d3ee`, `#34d399`, `#f87171`, `rgba(16,185,129,.05)`…). En modo claro varios **no cumplen AA** (p. ej. números de estadísticas `#34d399` sobre blanco ≈ 1,9:1).
5. **Emojis como iconos:** se ven distinto en cada sistema operativo y no siguen un estilo común entre pestañas.
6. **Feedback de estado fragmentado:** el punto de sincronización, el badge flotante "Guardando..." (que tapa el header) y el toast van por separado. Falta un estado "Guardado" y el error solo cambia el color, sin forma.
7. **Movimiento perpetuo:** `pulse` infinito en el punto de sync, `fireGlow` infinito en "TODO VALIDADO" y `alertPulse` en las alertas. Cansa la vista y no respeta `prefers-reduced-motion`.
8. **Glassmorphism sin fondo que desenfocar:** `backdrop-filter` en todas las tarjetas cuesta GPU en equipos modestos y no aporta nada visible.
9. **Detalles:** el overlay de carga es oscuro también en modo claro, `toast-info` tiene colores fijos de modo oscuro, `.btn-outline` se usa pero no existe, los inputs miden menos de 44 px en móvil y en móvil las pestañas desbordan sin ninguna señal.

---

## 2. Concepto visual

### A · "Clínico"
Superficies planas y blancas (o grafito en oscuro), bordes de 1 px, una sola tinta de acento azul institucional. La jerarquía la marcan la tipografía y el espacio, no el color. El movimiento es **confirmación**: aparece, confirma y se va.

### B · "Señal"
Conserva el violeta de la marca LabControl y le suma un **cian de "dato vivo"** para sincronización e información. Su motivo gráfico es una **línea de ECG** que aparece en la intro, en el estado "Sincronizado" y como subrayado de la pestaña activa. El movimiento es **respuesta**: cada acción tiene un pequeño rebote, un pulso o un trazo.

---

## 3. Paleta (contrastes medidos, WCAG 2.1)

### A · Clínico

| Token | Claro | Contraste* | Oscuro | Contraste* |
|---|---|---|---|---|
| `--bg` | `#F5F7FA` | — | `#0E1218` | — |
| `--surface` | `#FFFFFF` | — | `#161C25` | — |
| `--text` | `#17202E` | 15,3 | `#E6EAF0` | 15,6 |
| `--text-muted` | `#4A5568` | 7,0 | `#A3ADBB` | 8,3 |
| `--text-dim` | `#5F6B7C` | 5,0 | `#8A95A5` | 6,2 |
| `--accent` | `#2F5BD3` | 5,5 (blanco encima 5,9) | `#7EA2FF` | 7,6 (texto oscuro encima 7,6) |
| `--ok` (Agregados) | `#157F3B` | 4,7 | `#4CC38A` | 8,5 |
| `--bad` (Eliminados) | `#C0262D` | 5,5 | `#FF7A7A` | 7,4 |
| `--warn` | `#A45A00` | 4,8 | `#F2B54A` | 10,3 |

### B · Señal

| Token | Claro | Contraste* | Oscuro | Contraste* |
|---|---|---|---|---|
| `--bg` | `#F4F2FB` | — | `#0B0A14` | — |
| `--surface` | `#FFFFFF` | — | `#15131F` | — |
| `--text` | `#1A1530` | 15,9 | `#ECEAF6` | 16,6 |
| `--text-muted` | `#4E4766` | 7,8 | `#A9A3C4` | 8,2 |
| `--text-dim` | `#635C7E` | 5,6 | `#8C86A8` | 5,7 |
| `--accent` (marca) | `#6D3DF5` | 5,2 (blanco encima 5,8) | `#A68BFF` | 7,3 |
| `--signal` (dato vivo) | `#087489` | 4,9 | `#3DD6F0` | 11,3 |
| `--ok` | `#0F7A45` | 4,9 | `#3DDC97` | 11,1 |
| `--bad` | `#C42A3C` | 5,1 | `#FF6B7F` | 7,2 |
| `--warn` | `#A35800` | 4,8 | `#FFC15A` | 12,2 |

\* Contraste del texto sobre `--bg`. Todo el texto pasa AA (≥ 4,5:1). Los fondos tintados de Agregados y Eliminados se derivan con `color-mix()` al 8–14 % y se vuelven a verificar en la implementación.

**Regla semántica, en ambas opciones:** verde = Agregado / OK, rojo = Eliminado / error, ámbar = atención, acento = acción o navegación. El acento nunca se usa para un estado.

---

## 4. Tipografía

| | A · Clínico | B · Señal |
|---|---|---|
| Títulos | **IBM Plex Sans** 600 | **Bricolage Grotesque** 700 (con carácter, solo en títulos y cifras grandes) |
| Texto/UI | IBM Plex Sans 400/500 | Inter 400/500/600 |
| N° de petición, cifras | **IBM Plex Mono** 500 | Inter con `font-variant-numeric: tabular-nums` |

**Escala común** (rem, base 16 px). **Mínimo absoluto 12 px**, sin excepciones:

| Rol | Tamaño | Antes |
|---|---|---|
| Micro (metadatos, fechas) | 0,75 rem / 12 px | 8–9 px |
| Etiqueta / cabecera de tabla | 0,8125 rem / 13 px | 9–10 px |
| Cuerpo / inputs | 0,875 rem / 14 px | 12,5 px |
| Título de tarjeta | 1 rem / 16 px | 14 px |
| Título de panel | 1,125 rem / 18 px | 14 px |
| Cifra de estadística | 2,25 rem / 36 px | 29 px |

> Las clases `text-[8px]`, `text-[9px]` y `text-[10px]` vienen generadas desde `app.js`. Se re-mapean **por CSS** (sin tocar el JS) a la escala nueva dentro de `main`.

---

## 5. Sistema de espaciado y forma

- **Base de 4 px:** `--s-1: 4` · `--s-2: 8` · `--s-3: 12` · `--s-4: 16` · `--s-5: 24` · `--s-6: 32`.
- Padding de tarjeta: 16 px (móvil) / 20 px (escritorio). Separación entre tarjetas: 12 / 16 px.
- **Objetivo táctil:** inputs y botones de 36 px de alto en escritorio y **44 px en móvil** (`pointer: coarse`).
- **Radios:** A → 6 / 8 / 12 px · B → 10 / 14 / 20 px (más amable, "pill" en botones).
- **Elevación:** A → 1 px de borde + sombra mínima `0 1px 2px` · B → borde + sombra tintada de acento en hover/foco.
- **Foco visible:** anillo de 2 px de `--accent` con 2 px de separación en ambas opciones (hoy solo cambia el borde).

---

## 6. Iconografía

Un set SVG propio de trazo (grid de 20 px, trazo de 1,75 px, extremos redondeados) que sustituye los emojis **de `index.html`**:

| Pestaña | Icono |
|---|---|
| Errores | portapapeles con ± |
| Pizarra | matraz con nivel de líquido |
| Centros | edificio con cruz |
| Chat (oculto) | burbuja |
| Estadísticas | barras ascendentes |
| Admin | deslizadores |

También lleva iconos en los títulos de tarjeta (R10 ❄ → copo de trazo, C1, curvas, urgentes, recordatorios, cuadrantes) y en los estados del chip de sincronización.
En B, el icono de la pestaña activa hace un micro-trazo (`stroke-dashoffset`, 300 ms) al seleccionarse.

> Los emojis que genera `app.js` (🗑, ✕, ❌/✅ dentro de `<option>`, 🔥 TODO VALIDADO) **se mantienen**, salvo que autorices tocar esas cadenas (ver §9).

---

## 7. Principios de movimiento

- **P1. Solo anima lo que provocó el usuario.** El polling de 10 s re-renderiza con `innerHTML`, así que ninguna animación se engancha al render genérico. Las entradas se marcan solo tras una acción del usuario (añadir, guardar, cambiar de pestaña o de rango).
- **P2. Nunca bloquear.** Ninguna animación retrasa el foco, el `submit` ni la siguiente escritura. La lógica corre primero; la animación decora después y en paralelo.
- **P3. Determinista.** Las duraciones son fijas y no dependen de la red. Con sync lenta, la UI muestra "Guardando…" el tiempo que haga falta, pero sin animaciones nuevas.
- **P4. Sin movimiento perpetuo.** Se acaban los `pulse` infinitos. Los únicos bucles son el spinner de "Guardando…" (mientras dure) y nada más.
- **P5. `prefers-reduced-motion: reduce`.** Sin intro y sin desplazamientos ni escalas: solo cambios de opacidad ≤ 100 ms y los colores de estado. GSAP se desactiva (`gsap.globalTimeline.timeScale(1000)` + `set` directo).

**Tokens de movimiento**

| Token | A · Clínico | B · Señal |
|---|---|---|
| `--dur-1` (micro) | 120 ms | 150 ms |
| `--dur-2` (componente) | 180 ms | 240 ms |
| `--dur-3` (panel) | 240 ms | 400 ms |
| `--ease-out` | `cubic-bezier(.2,0,0,1)` | `cubic-bezier(.16,1,.3,1)` (expo-out) |
| `--ease-in` (salidas) | `cubic-bezier(.4,0,1,1)` | `cubic-bezier(.7,0,.84,0)` |
| `--ease-spring` | — (no se usa) | `cubic-bezier(.34,1.56,.64,1)` (back-out) |

---

## 8. Guion de movimiento por momento clave

### M1 · Carga inicial
| | A · Clínico | B · Señal |
|---|---|---|
| Secuencia | El logo y el wordmark hacen fade+rise de 8 px y luego la app aparece en fundido | 0–400 ms: el matraz se dibuja por trazo · 400–700 ms: el líquido sube y aparece el check · 700–1000 ms: el wordmark se revela letra a letra (stagger 25 ms) · 1000–1200 ms: la línea ECG cruza y el telón se abre hacia arriba |
| Duración | 600 ms | 1 200 ms (tope duro: 1 500 ms) |
| Easing | `--ease-out` | GSAP `expo.out` / `power2.inOut` en el telón |
| Omitible | Clic, tecla o toque la cierran al instante. Solo se muestra **una vez por sesión** (`sessionStorage`) | Igual |
| Si la red tarda | Tras la intro, las tarjetas de días y las tablas muestran **skeletons** con shimmer de 1,2 s (dentro de P4: se detiene al llegar los datos). El overlay de carga actual se re-estiliza como velo claro u oscuro según el tema, sin spinner genérico | Igual, con un shimmer en tono acento |
| Comunica | "Estás en LabControl y está listo." | "El sistema está vivo y conectado." |

### M2 · Cambio de pestaña
| | A · Clínico | B · Señal |
|---|---|---|
| Pestaña | Un indicador (fondo suave de acento) **se desliza** de la pestaña anterior a la nueva | El indicador se desliza con leve rebote y la línea ECG inferior se redibuja bajo la pestaña nueva; el icono hace un micro-trazo |
| Panel | Fade-in + rise de 4 px | Fade + rise de 8 px; las tarjetas entran con stagger de 30 ms (máx. 5 tarjetas y 150 ms en total) |
| Duración | 180 ms indicador · 160 ms panel | 240 ms indicador · 280 ms panel |
| Easing | `--ease-out` | `--ease-spring` en el indicador · `--ease-out` en el panel |
| Móvil | Desplaza automáticamente la pestaña activa al centro de la barra y muestra un degradado en los bordes para indicar que hay más | Igual |
| Comunica | "Te moviste a otra sección; el contenido es nuevo." | Ídem, con sensación de continuidad espacial |

### M3 · Añadir examen (Agregados / Eliminados)
| Evento | A · Clínico | B · Señal |
|---|---|---|
| Pill nueva | Escala 0,92→1 + fade · 140 ms | Escala 0,6→1 con rebote + flash del borde de la caja (verde o roja) · 260 ms · sonido "tic" opcional |
| Pill quitada | Un "fantasma" de la pill se desvanece y encoge en su lugar mientras la lista se recoloca · 120 ms | Fantasma que cae 6 px y se desvanece · 180 ms |
| Duplicado ("Ya está en la lista") | La pill existente parpadea su borde una vez · 200 ms | La pill existente hace un *shake* horizontal de 3 px × 3 · 240 ms |
| Técnica | Envolver `renderExamPills`: comparar con la lista anterior y marcar solo las nuevas con `.is-new`. El fantasma de salida es un clon absoluto, así que **la lista real ya se actualizó**; el foco no se mueve | Igual |
| Comunica | Verde = entra en Agregados, rojo = entra en Eliminados; "quedó anotado" | Igual, con más énfasis táctil |

### M4 · Guardar registro
| Resultado | A · Clínico | B · Señal |
|---|---|---|
| Éxito | El botón cambia su contenido por un ✓ dibujado (200 ms), se tiñe de verde 600 ms y vuelve. Las filas nuevas del día hacen un *highlight* de fondo verde que se apaga en 1,2 s | El botón hace un *morph*: el texto sale hacia arriba, entra el ✓ trazado, un anillo verde se expande desde el botón (400 ms) y vuelve. Las filas nuevas se deslizan desde la izquierda (stagger 40 ms) con *highlight*. Sonido de dos notas opcional |
| Error de validación | El botón hace un *shake* de 4 px (200 ms) y los campos vacíos (Día, Petición, Usuario o cajas de examen) muestran un borde rojo que se desvanece en 1,5 s | Igual, *shake* de 6 px (280 ms), borde rojo con un pulso único. Sonido grave opcional |
| Error de red (`apiPostBg` falla) | Toast de error + chip en estado Error (M5) | Igual |
| Duración total | ≤ 600 ms, **sin bloquear**: el foco ya está libre al instante | ≤ 800 ms, sin bloquear |
| Comunica | "Se guardó" / "Falta esto, aquí" | Igual |

**Toasts (ambas opciones):** se apilan (máx. 3) abajo a la derecha (en móvil, abajo al centro). Llevan icono y forma propios por tipo, entran con slide de 12 px + fade (180/240 ms), salen a los 3 s con fade de 150 ms y se pueden descartar con clic. El texto, ≤ 60 caracteres.

### M5 · Estado de sincronización (chip único en el header)
Fusiona el `#syncDot`/`#syncLabel` y el `#savingBadge` en **un solo chip**. Los IDs se mantienen; el chip lee su estado con los ganchos. Cada estado tiene **forma + color + icono**, así que se entiende sin leer y también con daltonismo.

| Estado | Forma/icono | A · Clínico | B · Señal |
|---|---|---|---|
| **Sincronizado** | círculo verde lleno | Estático | Un mini-trazo ECG se dibuja **una vez** tras cada poll exitoso (600 ms) y queda quieto |
| **Guardando…** | arco de acento girando | Spinner de 900 ms lineal mientras haya guardados pendientes | 3 barras tipo ecualizador (cian) |
| **Guardado** | ✓ verde trazado | El check se dibuja (240 ms), dura 1,5 s y pasa a Sincronizado con un fundido de 200 ms | Igual + leve escala 1→1,06→1 |
| **Sin conexión / error** | triángulo rojo con "!" | Aparece con fade y queda estático. El chip pasa a fondo rojo tenue | *Shake* único de 300 ms y luego estático |

Comunica, de un vistazo: verde quieto = todo bien · acento moviéndose = trabajando · check = terminó · rojo triangular = actuar.

### M6 · Cambio de tema claro/oscuro
| | A · Clínico | B · Señal |
|---|---|---|
| Transición | Fundido cruzado de toda la página (View Transitions API) | **Revelado circular** que nace del interruptor (View Transitions API, `clip-path`) |
| Duración | 200 ms | 450 ms |
| Easing | `ease-out` | `--ease-out` |
| Interruptor | Toggle con sol y luna; el pomo se desliza 180 ms | El sol se transforma en luna (los rayos se retraen y el disco se recorta) en 300 ms |
| Fallback | Navegadores sin View Transitions: los colores transicionan en 150 ms | Igual |
| Reduced motion | Cambio instantáneo | Cambio instantáneo |
| Comunica | "Cambió el entorno, el contenido es el mismo." | Ídem, con un gesto de "encender o apagar la luz" |

> La etiqueta de texto "🌙 Modo oscuro" se sustituye por el icono del propio toggle, con `aria-label` y `title`. El `#themeLabel` se mantiene en el DOM, visualmente oculto, para no romper `updateThemeLabel()`.

### M7 · Estadísticas (entrada animada)
| | A · Clínico | B · Señal |
|---|---|---|
| Contadores | Fade sin conteo (el número ya está) | Suben de 0 al valor (GSAP, 700 ms, `power2.out`) |
| Barras | Crecen con `scaleX` 0→1 desde la izquierda, 300 ms, stagger 20 ms | 600 ms, stagger 30 ms (máx. 15 por lista), `expo.out` |
| Colores | Agregados `--ok`, Eliminados `--bad`. Top usuarios: neutro `--text-muted` con barra roja **sólida** (sin degradado naranja) | Igual, con un degradado sutil del mismo tono |
| Cuándo | **Solo** al entrar en la pestaña o cambiar el rango. Los re-render del polling no re-animan (P1) | Igual |
| Comunica | Proporciones claras y comparables | La magnitud "se llena", con énfasis en el ranking |

### Extras de bajo impacto
- **"TODO VALIDADO":** A → insignia dorada estática con borde doble · B → brillo que recorre la insignia **una vez** al aparecer y luego descansa (se acaba el `fireGlow` infinito).
- **Alertas inferiores:** sin `alertPulse` infinito. Entran con slide-up, llevan barra lateral de color e icono, y se cierran con clic.
- **Modal de confirmación:** fade del fondo 150 ms + escala 0,96→1 en 180 ms. El botón peligroso, a la derecha y en rojo.
- **Sonido (solo B):** Web Audio API, sin archivos. Interruptor visible con altavoz en el header, **apagado por defecto** y guardado en `localStorage`. Volumen bajo (ganancia 0,04). Sonidos: tic (añadir), dos notas ascendentes (guardado), nota grave (error).

---

## 9. Decisiones que necesito que confirmes

1. **Capa de ganchos `ui/ui-hooks.js`:** envuelve funciones globales de `app.js` (llama a la original y solo añade clases o animaciones). No cambia la lógica, pero técnicamente intercepta funciones. ¿OK?
2. **Datos en pruebas locales:** la copia local usa la **API real**, así que probar "Guardar" escribiría en la planilla de producción. Propongo un modo `?demo=1`, **solo activo en `localhost`**, que simula la API con datos ficticios (`ui/demo-api.js`). Ese archivo no se integra a master. ¿OK?
3. **Textos de toasts en `app.js`:** algunos son confusos (p. ej. `"Petición"` a secas en Curvas y Urgentes → `"Ingresa N° de petición"`, `"✓"` en Cuadrantes → `"✓ Fila agregada"`). Cambiar **solo cadenas de texto** toca `app.js`. ¿Lo autorizas, o lo dejo fuera?
4. **Emojis dentro de `app.js`** (🗑, ✕, ❌/✅, 🔥): ¿se mantienen (por defecto) o autorizas cambiarlos por iconos o símbolos coherentes con el set nuevo?
5. **Tailwind Play CDN:** lo mantengo. Quitarlo exigiría reescribir todas las clases y es otro proyecto. Solo lo menciono como riesgo de rendimiento.

---

## 10. Plan de commits (rama `rediseno-visual`, ambas opciones en paralelo)

| # | Commit | Contenido |
|---|---|---|
| 0 | `style: base de tokens, tipografía, espaciado e iconos` | Fundación de A y B + selector `?look` + modo demo |
| 1 | `style(M1): intro y skeletons de carga` | |
| 2 | `style(M2): indicador y transición de pestañas` | |
| 3 | `style(M3): animación de pills de examen` | |
| 4 | `style(M4): feedback de guardado y toasts` | |
| 5 | `style(M5): chip de estado de sincronización` | |
| 6 | `style(M6): transición de tema` | |
| 7 | `style(M7): entrada animada de estadísticas` | |
| 8 | `style: pasada de refinamiento` | Tras las capturas de escritorio y móvil, en claro y oscuro |
| 9 | `chore: adoptar look elegido` | Borra el look descartado, el selector y el modo demo; merge a master y bump de cache (`?v=`) |
