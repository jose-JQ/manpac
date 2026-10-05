const COLS = [
  "Estado", "Fecha_Ejecucion", "Marca", "Modelo", "cantidad", "beneficiario", "ci",
  "costo_mas_alto", "costo_total", "marca_base", "modelo_base", "Tipo", "archivo",
];
const DETALLE = ["Marca", "Modelo", "cantidad", "beneficiario", "ci", "costo"];

const state = {
  files: [],
  corrida: [],
  tipos: ["SUV", "CAMION", "AUTOMOVIL", "CAMIONETA", "VAN", "BUS", "SCOOTER", "OTRO"],
  marcas: [],
  modelos: [],
  marcaSel: null,
  histEx: [],
  histRe: [],
  sel: null,
};

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

function esExito(reg) {
  const e = String((reg && (reg.Estado || reg.estado)) || "");
  return e.includes("ÉXITO") || e.toUpperCase().startsWith("EXITO");
}

function nombreDoc(reg) {
  return String((reg && (reg.archivo_guardado || reg.archivo)) || "").replace(/^.*[/\\]/, "");
}

function filaDe(reg) {
  return {
    Fecha_Ejecucion: reg.fecha_ejecucion || reg.Fecha_Ejecucion || "",
    Marca: reg.marca || reg.Marca || "",
    Modelo: reg.modelo || reg.Modelo || "",
    cantidad: reg.cantidad ?? 1,
    beneficiario: reg.beneficiario || "",
    ci: reg.ci || "",
    costo_mas_alto: reg.costo_mas_alto ?? 0,
    costo_total: reg.costo_total ?? 0,
    marca_base: reg.marca_base || "",
    modelo_base: reg.modelo_base || "",
    Tipo: reg.tipo || reg.Tipo || "OTRO",
    archivo: nombreDoc(reg),
    Estado: reg.Estado || (esExito(reg) ? "EXITO" : String(reg.estado || "REVISION")),
    _reg: reg,
  };
}

async function health() {
  try {
    const r = await fetch("/api/health");
    const data = r.ok ? await r.json() : {};
    $("#health").classList.toggle("on", r.ok);
    $("#health").classList.toggle("off", !r.ok);
    $("#health").title = data.aviso || data.ollama_host || "";
  } catch {
    $("#health").classList.add("off");
  }
}

function setView(name) {
  $$(".nav-btn").forEach((b) => b.classList.toggle("active", b.dataset.view === name));
  $$(".view").forEach((v) => v.classList.add("hidden"));
  $(`#view-${name}`).classList.remove("hidden");
  const copy = {
    procesar: ["Procesar documentos", "Selecciona una fila para editar a la derecha y ver el PDF."],
    reportes: ["Reportes", "Corrida reciente e histórico. El estado se recalcula al guardar, no se escribe a mano."],
    vehiculos: ["Catálogo de vehículos", "Marcas y modelos usados para validar las extracciones."],
  }[name];
  $("#page-title").textContent = copy[0];
  $("#page-sub").textContent = copy[1];
  cerrarMenu();
  if (name === "vehiculos") cerrarPanel();
  if (name === "reportes") loadCorrida();
  if (name === "vehiculos") loadCatalogo();
}

function claveFila(row) {
  const reg = row._reg || row;
  return nombreDoc(reg) || `${row.Marca}|${row.Modelo}|${row.ci}|${row.Fecha_Ejecucion}`;
}

function renderTable(el, rows, opts = {}) {
  if (!rows.length) {
    el.innerHTML = `<div class="empty">${opts.empty || "Sin registros."}</div>`;
    return;
  }
  const prevWrap = el.querySelector(".table-wrap");
  const sl = prevWrap ? prevWrap.scrollLeft : 0;
  const st = prevWrap ? prevWrap.scrollTop : 0;
  const cols = opts.cols || COLS;
  const head = cols.map((c) => `<th>${c}</th>`).join("");
  const selKey = state.sel && state.sel.key;
  const body = rows.map((row, i) => {
    const key = claveFila(row);
    const cells = cols.map((c) => {
      if (c === "Estado") {
        const ok = String(row[c] || "").toUpperCase().startsWith("EXITO");
        return `<td><span class="badge ${ok ? "ok" : "warn"}">${esc(row[c] || "")}</span></td>`;
      }
      return `<td title="${esc(row[c])}">${esc(row[c])}</td>`;
    }).join("");
    const sel = selKey && key === selKey ? " selected" : "";
    return `<tr class="${sel.trim()}" data-row="${i}" data-key="${esc(key)}">${cells}</tr>
      <tr class="items-sub hidden" data-sub="${i}"><td colspan="${cols.length}"></td></tr>`;
  }).join("");
  el.innerHTML = `<div class="table-wrap"><table class="data"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  const wrap = el.querySelector(".table-wrap");
  if (wrap) {
    wrap.scrollLeft = sl;
    wrap.scrollTop = st;
  }
  el.querySelectorAll("tr[data-row]").forEach((tr) => {
    tr.addEventListener("click", () => {
      const i = +tr.dataset.row;
      abrirPanel(rows[i], opts.editable !== false);
    });
  });
}

function marcarFilaSeleccionada() {
  const key = state.sel && state.sel.key;
  $$("tr[data-row]").forEach((tr) => {
    tr.classList.toggle("selected", Boolean(key) && tr.dataset.key === key);
  });
}

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function fillTipos(valor) {
  const sel = $("#panel-tipo");
  sel.innerHTML = state.tipos.map((t) => `<option ${t === valor ? "selected" : ""}>${t}</option>`).join("");
}

function itemsDe(reg) {
  return (reg && reg.items) || [];
}

function pintarItems(items) {
  const box = $("#panel-items");
  if (!items.length) {
    box.innerHTML = `<div class="empty">Sin líneas de detalle.</div>`;
    return;
  }
  box.innerHTML = items.map((it, i) => `
    <div class="item-row" data-item="${i}">
      <input data-k="Marca" placeholder="Marca" value="${esc(it.Marca || "")}">
      <input data-k="Modelo" placeholder="Modelo" value="${esc(it.Modelo || it.Descripcion || "")}">
      <input data-k="Cantidad" placeholder="Cant." value="${esc(it.Cantidad ?? it.cantidad ?? 1)}">
      <input data-k="Valor" placeholder="Costo" value="${esc(it.Valor ?? it.costo ?? 0)}">
      <button type="button" class="icon-btn" data-del-item="${i}" aria-label="Quitar">×</button>
    </div>`).join("");
  box.querySelectorAll("[data-del-item]").forEach((b) => {
    b.addEventListener("click", (e) => {
      e.preventDefault();
      const list = leerItems();
      list.splice(+b.dataset.delItem, 1);
      pintarItems(list);
    });
  });
}

function leerItems() {
  return [...$("#panel-items").querySelectorAll(".item-row")].map((row) => {
    const get = (k) => row.querySelector(`[data-k="${k}"]`).value;
    return {
      Marca: get("Marca"),
      Modelo: get("Modelo"),
      Cantidad: Number(get("Cantidad") || 1),
      Valor: Number(get("Valor") || 0),
      costo: Number(get("Valor") || 0),
    };
  });
}

function mostrarDocumento(nombre) {
  const pdf = $("#panel-pdf");
  const img = $("#panel-img");
  if (!nombre) {
    pdf.removeAttribute("src");
    img.removeAttribute("src");
    img.classList.add("hidden");
    return;
  }
  const url = `/api/documento?nombre=${encodeURIComponent(nombre)}`;
  const ext = (nombre.split(".").pop() || "").toLowerCase();
  if (["png", "jpg", "jpeg", "webp"].includes(ext)) {
    pdf.removeAttribute("src");
    img.src = url;
    img.classList.remove("hidden");
    pdf.classList.add("hidden");
  } else {
    img.classList.add("hidden");
    pdf.classList.remove("hidden");
    img.removeAttribute("src");
    pdf.src = url + (url.includes("?") ? "&" : "?") + "t=" + Date.now();
  }
}

function indiceCorrida(reg) {
  const nom = nombreDoc(reg);
  if (nom) {
    const i = state.corrida.findIndex((r) => nombreDoc(r) === nom);
    if (i >= 0) return i;
  }
  return state.corrida.indexOf(reg);
}

function abrirPanel(row, editable = true) {
  const main = $("main");
  const mainTop = main ? main.scrollTop : 0;
  const reg = (row && row._reg) || row || {};
  const idx = indiceCorrida(reg);
  const puedeEditar = editable && idx >= 0;
  state.sel = { key: claveFila(row), editable: puedeEditar, indice: idx };
  $("#panel-edicion").classList.add("open");
  $("#panel-empty").classList.add("hidden");
  $("#panel-body").classList.remove("hidden");
  $("#panel-archivo").textContent = nombreDoc(reg) || row.archivo || "Documento";
  const ok = esExito(reg);
  const badge = $("#panel-estado");
  badge.textContent = row.Estado || reg.Estado || reg.estado || "";
  badge.classList.toggle("ok", ok);
  badge.classList.toggle("warn", !ok);
  const form = $("#panel-form");
  form.marca.value = reg.marca || row.Marca || "";
  form.modelo.value = reg.modelo || row.Modelo || "";
  form.cantidad.value = reg.cantidad ?? row.cantidad ?? 1;
  form.beneficiario.value = reg.beneficiario || "";
  form.ci.value = reg.ci || "";
  form.costo_mas_alto.value = reg.costo_mas_alto ?? row.costo_mas_alto ?? 0;
  form.costo_total.value = reg.costo_total ?? row.costo_total ?? 0;
  fillTipos(reg.tipo || row.Tipo || "OTRO");
  pintarItems(itemsDe(reg));
  $("#btn-panel-guardar").disabled = !puedeEditar;
  $("#panel-msg").textContent = puedeEditar
    ? ""
    : "Solo lectura: este registro no está en la corrida reciente. El estado no se edita a mano.";
  mostrarDocumento(nombreDoc(reg));
  marcarFilaSeleccionada();
  if (main) main.scrollTop = mainTop;
}

function abrirMenu() {
  document.body.classList.add("nav-open");
  $("#btn-menu")?.setAttribute("aria-expanded", "true");
  $("#btn-menu")?.setAttribute("aria-label", "Cerrar menú");
}

function cerrarMenu() {
  document.body.classList.remove("nav-open");
  $("#btn-menu")?.setAttribute("aria-expanded", "false");
  $("#btn-menu")?.setAttribute("aria-label", "Abrir menú");
}

function toggleMenu() {
  if (document.body.classList.contains("nav-open")) cerrarMenu();
  else abrirMenu();
}

function cerrarPanel() {
  state.sel = null;
  $("#panel-edicion").classList.remove("open");
  $("#panel-body").classList.add("hidden");
  $("#panel-empty").classList.remove("hidden");
  mostrarDocumento("");
  marcarFilaSeleccionada();
}

function pintarTablas() {
  const live = $("#live-table");
  const exEl = $("#tabla-exitos");
  const revEl = $("#tabla-revision");
  if (live) {
    renderTable(live, state.corrida.map(filaDe), { empty: "Todavía no hay resultados en esta sesión." });
  }
  if (exEl && revEl) {
    const filas = state.corrida.map(filaDe);
    renderTable(exEl, filas.filter((f) => esExito(f._reg)), { empty: "No hay éxitos en la corrida reciente." });
    renderTable(revEl, filas.filter((f) => !esExito(f._reg)), { empty: "No hay revisiones en la corrida reciente." });
  }
  if (state.histEx.length || state.histRe.length) pintarHistorico();
}

function pintarHistorico() {
  const exEl = $("#tabla-hist-ex");
  const revEl = $("#tabla-hist-re");
  if (!exEl || !revEl) return;
  renderTable(exEl, state.histEx.map(filaDe), { empty: "Histórico de éxitos vacío." });
  renderTable(revEl, state.histRe.map(filaDe), { empty: "Histórico de revisión vacío." });
}

async function loadCorrida() {
  const r = await fetch("/api/corrida");
  const data = await r.json();
  state.corrida = data.registros || [];
  if (data.tipos) state.tipos = data.tipos;
  pintarTablas();
}

async function loadHistorico() {
  const r = await fetch("/api/historico");
  const data = await r.json();
  state.histEx = data.exitos || [];
  state.histRe = data.revision || [];
  pintarHistorico();
}

async function loadCatalogo() {
  const r = await fetch("/api/catalogo");
  const data = await r.json();
  state.marcas = data.marcas || [];
  state.modelos = data.modelos || [];
  if (data.tipos) state.tipos = data.tipos;
  pintarMarcas();
}

function pintarMarcas() {
  const box = $("#lista-marcas");
  box.innerHTML = state.marcas.map((m) =>
    `<button type="button" data-id="${m.id_marca}" class="${state.marcaSel === m.id_marca ? "active" : ""}">${esc(m.nombre)}</button>`
  ).join("") || `<div class="empty">Sin marcas.</div>`;
  box.querySelectorAll("button").forEach((b) => {
    b.addEventListener("click", () => {
      state.marcaSel = +b.dataset.id;
      pintarMarcas();
      pintarModelos();
    });
  });
  pintarModelos();
}

function pintarModelos() {
  const marca = state.marcas.find((m) => m.id_marca === state.marcaSel);
  $("#modelos-title").textContent = marca ? `Modelos de ${marca.nombre}` : "Modelos";
  const rows = state.modelos.filter((m) => !state.marcaSel || m.id_marca === state.marcaSel);
  if (!rows.length) {
    $("#tabla-modelos").innerHTML = `<div class="empty">Selecciona una marca.</div>`;
    return;
  }
  $("#tabla-modelos").innerHTML = `<div class="table-wrap"><table class="data">
    <thead><tr><th>Modelo</th><th>Carrocería</th><th></th></tr></thead>
    <tbody>${rows.map((m) => `<tr>
      <td>${esc(m.nombre)}</td><td>${esc(m.tipo_carroceria)}</td>
      <td class="row-actions">
        <button type="button" class="icon-btn" data-edit="${m.id_modelo}">✎</button>
        <button type="button" class="icon-btn" data-del="${m.id_modelo}">✕</button>
      </td>
    </tr>`).join("")}</tbody></table></div>`;
  $("#tabla-modelos").querySelectorAll("tbody tr").forEach((tr) => {
    tr.style.cursor = "default";
  });
  $("#tabla-modelos").querySelectorAll("[data-edit]").forEach((b) => {
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      editarModelo(+b.dataset.edit);
    });
  });
  $("#tabla-modelos").querySelectorAll("[data-del]").forEach((b) => {
    b.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!confirm("¿Borrar este modelo?")) return;
      await fetch(`/api/catalogo/modelos/${b.dataset.del}`, { method: "DELETE" });
      loadCatalogo();
    });
  });
}

function dialogForm(title, fields) {
  return new Promise((resolve) => {
    $("#dlg-title").textContent = title;
    $("#dlg-body").innerHTML = fields.map((f) => {
      if (f.type === "select") {
        return `<label>${f.label}<select name="${f.name}">${f.options.map((o) =>
          `<option ${o === f.value ? "selected" : ""}>${o}</option>`).join("")}</select></label>`;
      }
      return `<label>${f.label}<input name="${f.name}" value="${esc(f.value || "")}"></label>`;
    }).join("");
    const dlg = $("#dlg");
    let done = false;
    const finish = (val) => {
      if (done) return;
      done = true;
      dlg.removeEventListener("close", onClose);
      resolve(val);
    };
    const onClose = () => {
      if (dlg.returnValue !== "ok") return finish(null);
      const fd = new FormData($("#dlg-form"));
      const out = {};
      fields.forEach((f) => { out[f.name] = fd.get(f.name); });
      finish(out);
    };
    $("#dlg-cancel").onclick = () => dlg.close("cancel");
    dlg.addEventListener("close", onClose);
    dlg.showModal();
  });
}

function bindFiles() {
  const input = $("#file-input");
  const drop = $("#dropzone");
  const list = $("#file-list");
  const sync = () => {
    state.files = [...(input.files || [])];
    list.innerHTML = state.files.map((f) => `<li>${esc(f.name)}</li>`).join("");
    $("#btn-procesar").disabled = !state.files.length;
  };
  input.addEventListener("change", sync);
  ["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => {
    e.preventDefault(); drop.classList.add("drag");
  }));
  ["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => {
    e.preventDefault(); drop.classList.remove("drag");
  }));
  drop.addEventListener("drop", (e) => {
    input.files = e.dataTransfer.files;
    sync();
  });
}

async function procesar() {
  const files = state.files;
  if (!files.length) return;
  $("#btn-procesar").disabled = true;
  $("#prog-wrap").classList.remove("hidden");
  await fetch("/api/corrida/iniciar", { method: "POST" });
  state.corrida = [];
  pintarTablas();
  try {
    for (let i = 0; i < files.length; i += 1) {
      $("#proc-status").textContent = `Procesando ${files[i].name} (${i + 1}/${files.length})`;
      $("#prog-bar").style.width = `${((i) / files.length) * 100}%`;
      const fd = new FormData();
      fd.append("file", files[i]);
      const res = await fetch("/api/procesar-factura", { method: "POST", body: fd });
      const reader = res.body.getReader();
      const dec = new TextDecoder();
      let buf = "";
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += dec.decode(value, { stream: true });
        const lines = buf.split("\n");
        buf = lines.pop();
        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            state.corrida.push(JSON.parse(line));
            pintarTablas();
          } catch { /* línea incompleta */ }
        }
      }
      if (buf.trim()) {
        try { state.corrida.push(JSON.parse(buf)); pintarTablas(); } catch { /* fin */ }
      }
    }
    $("#prog-bar").style.width = "100%";
    $("#proc-status").textContent = "Lote terminado. Selecciona una fila para editar y ver el PDF.";
  } finally {
    await fetch("/api/corrida/finalizar", { method: "POST" });
    await loadCorrida();
    $("#btn-procesar").disabled = false;
  }
}

async function guardarPanel() {
  if (!state.sel || state.sel.indice < 0) return;
  const form = $("#panel-form");
  $("#btn-panel-guardar").disabled = true;
  $("#panel-msg").textContent = "Revalidando…";
  try {
    const r = await fetch("/api/corrida/validar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        indice: state.sel.indice,
        archivo_guardado: $("#panel-archivo").textContent,
        Marca: form.marca.value,
        Modelo: form.modelo.value,
        cantidad: form.cantidad.value,
        beneficiario: form.beneficiario.value,
        ci: form.ci.value,
        costo_mas_alto: form.costo_mas_alto.value,
        costo_total: form.costo_total.value,
        Tipo: form.tipo.value,
        items: leerItems(),
      }),
    });
    const data = await r.json();
    if (!r.ok) {
      $("#panel-msg").textContent = data.detail || "No se pudo revalidar.";
      return;
    }
    await loadCorrida();
    const idx = typeof data.indice === "number" ? data.indice : state.sel.indice;
    const reg = state.corrida[idx] || data.registro;
    if (reg) abrirPanel(filaDe(reg), true);
    $("#panel-msg").textContent = esExito(reg) ? "Pasó a Éxitos." : "Quedó en Revisión.";
  } catch (e) {
    $("#panel-msg").textContent = String(e);
  } finally {
    $("#btn-panel-guardar").disabled = false;
  }
}

async function guardarYAceptar() {
  const r = await fetch("/api/corrida/aceptar", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: "{}",
  });
  const data = await r.json();
  $("#oracle-msg").textContent = data.mensaje || "";
}

function editarModelo(id) {
  const m = state.modelos.find((x) => x.id_modelo === id);
  dialogForm("Editar modelo", [
    { name: "nombre", label: "Modelo", value: m.nombre },
    { name: "tipo_carroceria", label: "Carrocería", type: "select", options: state.tipos, value: m.tipo_carroceria },
  ]).then(async (f) => {
    if (!f) return;
    await fetch(`/api/catalogo/modelos/${id}`, {
      method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(f),
    });
    loadCatalogo();
  });
}

function aplicarAnchoPanel(px) {
  const min = 280;
  const max = Math.max(min, Math.min(820, Math.floor(window.innerWidth * 0.72)));
  const w = Math.round(Math.max(min, Math.min(max, Number(px) || 400)));
  const val = `${w}px`;
  document.documentElement.style.setProperty("--panel-w", val);
  const panel = $("#panel-edicion");
  if (panel) panel.style.setProperty("--panel-w", val);
  try { localStorage.setItem("manpac-panel-w", String(w)); } catch { /* ignore */ }
  return w;
}

function bindPanelResize() {
  const handle = $("#panel-resizer");
  const panel = $("#panel-edicion");
  if (!handle || !panel) return;
  let startX = 0;
  let startW = 0;
  const saved = Number(localStorage.getItem("manpac-panel-w") || 0);
  if (saved >= 280) aplicarAnchoPanel(saved);

  handle.addEventListener("pointerdown", (e) => {
    if (window.matchMedia("(max-width: 960px)").matches) return;
    e.preventDefault();
    startX = e.clientX;
    startW = panel.getBoundingClientRect().width;
    handle.setPointerCapture(e.pointerId);
    document.body.classList.add("panel-resizing");
  });
  handle.addEventListener("pointermove", (e) => {
    if (!document.body.classList.contains("panel-resizing")) return;
    aplicarAnchoPanel(startW + (startX - e.clientX));
  });
  const stop = () => document.body.classList.remove("panel-resizing");
  handle.addEventListener("pointerup", stop);
  handle.addEventListener("pointercancel", stop);
  handle.addEventListener("dblclick", () => aplicarAnchoPanel(400));
  handle.addEventListener("keydown", (e) => {
    const actual = panel.getBoundingClientRect().width;
    if (e.key === "ArrowLeft") {
      e.preventDefault();
      aplicarAnchoPanel(actual + 24);
    }
    if (e.key === "ArrowRight") {
      e.preventDefault();
      aplicarAnchoPanel(actual - 24);
    }
  });
}

function init() {
  $$(".nav-btn").forEach((b) => b.addEventListener("click", () => setView(b.dataset.view)));
  $$(".tab").forEach((b) => b.addEventListener("click", () => {
    $$(".tab").forEach((t) => t.classList.toggle("active", t === b));
    const hist = b.dataset.rep === "historico";
    $("#rep-reciente").classList.toggle("hidden", hist);
    $("#rep-historico").classList.toggle("hidden", !hist);
    if (hist) loadHistorico();
  }));
  $("#panel-cerrar").addEventListener("click", cerrarPanel);
  bindPanelResize();
  $("#btn-menu").addEventListener("click", toggleMenu);
  $("#nav-backdrop").addEventListener("click", cerrarMenu);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") cerrarMenu();
  });
  $("#btn-item-add").addEventListener("click", () => {
    const list = leerItems();
    list.push({ Marca: "", Modelo: "", Cantidad: 1, Valor: 0 });
    pintarItems(list);
  });
  $("#btn-panel-guardar").addEventListener("click", guardarPanel);
  $("#btn-procesar").addEventListener("click", procesar);
  $("#btn-aceptar").addEventListener("click", guardarYAceptar);
  $("#btn-add-marca").addEventListener("click", async () => {
    const f = await dialogForm("Nueva marca", [
      { name: "nombre", label: "Nombre" },
      { name: "origen", label: "Origen" },
    ]);
    if (!f) return;
    await fetch("/api/catalogo/marcas", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(f) });
    loadCatalogo();
  });
  $("#btn-add-modelo").addEventListener("click", async () => {
    const marca = state.marcas.find((m) => m.id_marca === state.marcaSel);
    if (!marca) { alert("Elige una marca."); return; }
    const f = await dialogForm("Nuevo modelo", [
      { name: "nombre", label: "Modelo" },
      { name: "tipo_carroceria", label: "Carrocería", type: "select", options: state.tipos, value: "OTRO" },
    ]);
    if (!f) return;
    await fetch("/api/catalogo/modelos", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ marca: marca.nombre, nombre: f.nombre, tipo_carroceria: f.tipo_carroceria }),
    });
    loadCatalogo();
  });
  bindFiles();
  health();
  setInterval(health, 15000);
  loadCorrida();
}

init();
