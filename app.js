/* =============================
   app.js — Callcenter Musicala
   Paginación + búsqueda global + modal editable
   ============================= */

'use strict';

(async function () {
  // ------- DOM -------
  const statusEl   = document.getElementById('status');
  const table      = document.getElementById('dataTable');
  const sheetSel   = document.getElementById('sheetSelect');
  const searchBox  = document.getElementById('search');
  const btnNew     = document.getElementById('btnNew');

  // Paginación
  const pagerEl    = document.getElementById('pager');
  const btnPrev    = document.getElementById('btnPrev');
  const btnNext    = document.getElementById('btnNext');
  const pageInfo   = document.getElementById('pageInfo');
  const pageSizeEl = document.getElementById('pageSize');

  // Modal
  const modal         = document.getElementById('clientModal');
  const modalBody     = document.getElementById('clientBody');
  const modalTitle    = document.getElementById('clientTitle');
  const modalSubtitle = document.getElementById('clientSubtitle');

  // 🔔 Avisos de urgentes/agendados (en la página)
  const alertsEl      = document.getElementById('alerts');

  // ------- Config -------
  if (!window.API_BASE) throw new Error('No se encontró window.API_BASE. Define la URL /exec en index.html');
  const API_BASE = window.API_BASE;

  // 🔒 Clave SOLO "ID"
  const KEY_CANDIDATES = ['ID'];

  // ====== Catálogos (listas fijas) ======
  const OPTIONS = {
    Grupo: ["Musiadultos","Musibabies","Musicalitos","Musigrandes","Musikids","Musiteens"],
    ARTE: ["Música","Baile","Artes plásticas","Teatro","Vacacionales","Todos"],
    Modalidad: [
      "Domicilio","Hogar y virtual","Sede","Sede,Virtual y Domicilio",
      "Sede y Hogar","Virtual","Hogar","Plataforma Online","Sede y Virtual"
    ],
    "Curso/Plan": [
      "Personalizado","Grupal","Hogar","Virtual","Vacacional",
      "AutoMusicala","MusiGym","Curso de Formación","Preuniversitario",
      "Grupal y personalizado","Musifamiliar"
    ],
    Listado: [
      "Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto",
      "Septiembre","Octubre","Noviembre","Diciembre","2023","2024","Bloqueados",
      "Sede","Activo","Sin agendar","Clase de  Prueba","FESICOL","Vacacionales",
      "Virtual","A hogar","No interesad@","Otros horarios","Distancia","Inscripción","No disponible"
    ],
    Listado1: [
      "Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto",
      "Septiembre","Octubre","Noviembre","Diciembre","2023","2024","Bloqueados",
      "Sede","Activo","Sin agendar","Clase de  Prueba","FESICOL","Vacacionales",
      "Virtual","A hogar","No interesad@","Otros horarios","Distancia","Inscripción","No disponible"
    ],
    Asesor: ["Alek Caballero","Catalina Medina","Camila Rodríguez"],
    "Canal de comunicación": ["Llamada","WhatsApp","Keybe","Instagram","Facebook","Wix","TikTok"],
    Prioridad: ["Alta","Media","Baja"]
  };

  // ====== Dependencias Arte/Instrumento ======
  const ART_TO_OPTIONS = {
    "Música": [
      "Piano","Guitarra","Canto","Violín","Batería","Cello","Bajo eléctrico","Ukelele",
      "Acordeón","Bandola","Iniciación Musical","Teoría","Composición música",
      "Flauta Dulce","Saxofón","Requinto","Trombón","Trompeta","Percusión","Xilófono","Jazz"
    ],
    "Baile": [
      "Bailes Latinos","Danza clásica y contemporánea","Bailes Urbanos","Danzas folclóricas",
      "Ballet","Salsa","Bongo"
    ],
    "Artes plásticas": [
      "Dibujo","Pintura","Óleo","Plastilina","Técnicas mixtas","Exploración Artística","Manualidades"
    ],
    "Teatro": ["Teatro"],
    "Vacacionales": ["Vacacionales"],
    "Todos": ["Piano","Guitarra","Canto","Violín","Batería","Bailes Latinos","Dibujo","Pintura","Teatro","Vacacionales"]
  };

  const ARTE_COLS = ["Arte I","Arte II","Arte III"];
  const INSTRUMENT_COLS = [
    "Instrumento/Estilo/Técnica I",
    "Instrumento/Estilo/Técnica II",
    "Instrumento/Estilo/Técnica III"
  ];

  // ------- Estado -------
  let currentSheet   = '';
  let limit          = parseInt(pageSizeEl?.value || '200', 10);
  let offset         = 0;
  let total          = 0;
  let currentHeaders = [];
  let currentRows    = [];
  let originalRow    = null;
  let keyColumnInUse = '';
  let creatingNew    = false;
  // Orden de columnas (sorting)
  let sortState = { key: null, dir: 1 }; // dir: 1 = asc, -1 = desc

  // ===== Búsqueda GLOBAL =====
  const allRowsCache = new Map(); // sheet -> { headers, rows, total, ts }
  let searchActive   = false;
  let displayedRows  = [];        // filas actualmente renderizadas (página o filtradas)

  /* =============================
     1) Cargar hojas + primera página
     ============================= */
  const meta = await fetchJSON(`${API_BASE}?mode=meta&_ts=${Date.now()}`).catch(showError);
  if (!meta || !Array.isArray(meta.sheets) || meta.sheets.length === 0) {
    status('No se encontraron hojas.');
    return;
  }
  sheetSel.innerHTML = meta.sheets
    .map((n) => `<option value="${escAttr(n)}">${esc(n)}</option>`)
    .join('');

  const initial = meta.sheets.find((n) => /base de datos general/i.test(n)) || meta.sheets[0];
  sheetSel.value = initial;
  currentSheet   = initial;

  await loadPage();

  /* =============================
     2) Eventos UI
     ============================= */
  sheetSel.addEventListener('change', async (e) => {
    currentSheet = e.target.value;
    offset = 0;
    searchBox.value = '';
    searchActive = false;
    pagerEl?.classList.remove('hidden');
    await loadPage();
  });

  btnPrev?.addEventListener('click', async () => {
    if (searchActive) return; // no mover paginación durante búsqueda
    if (offset <= 0) return;
    offset = Math.max(0, offset - limit);
    await loadPage();
  });

  btnNext?.addEventListener('click', async () => {
    if (searchActive) return;
    if (offset + limit >= total) return;
    offset = offset + limit;
    await loadPage();
  });

  pageSizeEl?.addEventListener('change', async () => {
    if (searchActive) return;
    limit = parseInt(pageSizeEl.value, 10) || 200;
    offset = 0;
    await loadPage();
  });

  // Búsqueda GLOBAL con debounce
  let debounce;
  searchBox.addEventListener('input', () => {
    clearTimeout(debounce);
    debounce = setTimeout(applySearch, 250);
  });

  // Nuevo cliente
  btnNew?.addEventListener('click', () => openNewModal());

  // Delegación: click en el "Nombre" abre modal de edición
  table.addEventListener('click', (e) => {
    const btn = e.target.closest('.name-link');
    if (!btn) return;
    const idx = Number(btn.dataset.rowindex);
    const row = displayedRows[idx]; // usa las filas visibles (página o búsqueda)
    if (row) openModal(row);
  });

  // Cerrar modal
  modal.addEventListener('click', (e) => {
    if (e.target.dataset.close === '1') closeModal();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
  });

  /* =============================
     3) Lógica de datos (paginado normal)
     ============================= */
  async function loadPage() {
    status(`Cargando “${currentSheet}”…`);
    const url = `${API_BASE}?mode=data&sheet=${encodeURIComponent(currentSheet)}&limit=${limit}&offset=${offset}&_ts=${Date.now()}`;
    const res = await fetchJSON(url).catch(showError);
    if (!res || !Array.isArray(res.headers)) {
      status('Error cargando datos.');
      return;
    }

    currentHeaders = res.headers;
    currentRows    = res.rows || [];
    total          = parseInt(res.total || 0, 10);

    keyColumnInUse = pickKeyColumn(currentHeaders);
    if (!keyColumnInUse) {
      status('⚠️ Esta hoja no tiene columna "ID". Verás datos, pero no podrás guardar cambios.');
    } else {
      status(`Listo. Clave: “${keyColumnInUse}”. Mostrando ${currentRows.length.toLocaleString()} de ${total.toLocaleString()} registros.`);
    }

    displayedRows = currentRows;
    renderTable(currentHeaders, displayedRows);

    const page  = Math.floor(offset / limit) + 1;
    const pages = Math.max(1, Math.ceil(total / limit));
    if (pageInfo) {
      pageInfo.textContent = `Página ${page.toLocaleString()} de ${pages.toLocaleString()} — ${total.toLocaleString()} registros`;
    }
    if (btnPrev) btnPrev.disabled = offset <= 0;
    if (btnNext) btnNext.disabled = offset + limit >= total;

    // 🔔 Calcular y pintar alertas
    await computeAndRenderAlerts();
  }

  function renderTable(headers, rows) {
  const nameIdx = headers.findIndex((h) => h.trim().toLowerCase() === 'nombre');

  // Construye THEAD con headers “clicables” para ordenar
  const thead = `<thead><tr>${
    headers.map((h) => {
      const isActive = sortState.key === h;
      const dirAttr  = isActive ? ` data-dir="${sortState.dir}"` : '';
      return `<th class="th-sortable" data-key="${escAttr(h)}"${dirAttr}>
                <span>${esc(h)}</span>
                <span class="th-sort-ind"></span>
              </th>`;
    }).join('')
  }</tr></thead>`;

  // Si hay una columna activa, ordena una copia de las filas visibles
  let toRender = rows;
  if (sortState.key) {
    const key = sortState.key;
    const dir = sortState.dir;
    toRender = [...rows].map((r, i) => ({ r, i })) // estable para empates
      .sort((a, b) => {
        const cmp = compareByKey(a.r, b.r, key);
        return (cmp !== 0) ? (cmp * dir) : (a.i - b.i);
      })
      .map(x => x.r);
  }

  const tbody = `<tbody>${
    toRender.map((r, i) => {
      const tds = headers.map((h, colIdx) => {
        const val = r[h];
        if (colIdx === nameIdx) {
          const label = (val == null || String(val).trim() === '') ? '(Sin nombre)' : String(val);
          return `<td><button class="linkish name-link" data-rowindex="${i}">${esc(label)}</button></td>`;
        }
        return `<td>${pretty(val)}</td>`;
      }).join('');
      return `<tr>${tds}</tr>`;
    }).join('')
  }</tbody>`;

  table.innerHTML = thead + tbody;

  // fuerza ancho mínimo por cantidad de columnas (para scroll horizontal)
  const pxPerCol = 140;
  table.style.minWidth = (headers.length * pxPerCol) + 'px';

  // Listeners para ordenar al hacer clic en el TH
  table.querySelectorAll('thead th.th-sortable').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.getAttribute('data-key') || '';
      if (!key) return;
      if (sortState.key === key) {
        // alterna asc/desc
        sortState.dir = sortState.dir === 1 ? -1 : 1;
      } else {
        sortState.key = key;
        sortState.dir = 1;
      }
      // re-render con el nuevo estado de orden
      renderTable(headers, rows);
    });
  });
}

  // Carga TODA la hoja usando paginación del API y la junta
  async function loadAllRowsForSheet(sheetName) {
    const headers = currentHeaders.length ? currentHeaders : (await fetchJSON(`${API_BASE}?mode=data&sheet=${encodeURIComponent(sheetName)}&limit=1&offset=0&_ts=${Date.now()}`)).headers;
    const pageLimit = 1000; // tamaño grande para acelerar
    let off = 0;
    let tot = 0;
    let all = [];

    while (true) {
      const url = `${API_BASE}?mode=data&sheet=${encodeURIComponent(sheetName)}&limit=${pageLimit}&offset=${off}&_ts=${Date.now()}`;
      const res = await fetchJSON(url);
      const rows = res.rows || [];
      tot = parseInt(res.total || rows.length, 10);
      all = all.concat(rows);
      off += pageLimit;
      if (all.length >= tot || rows.length === 0) break;
    }

    return { headers, rows: all, total: tot, ts: Date.now() };
  }

  /* =============================
     5) Modales
     ============================= */

  // Abre modal para fila existente (editar)
  function openModal(row) {
    creatingNew = false;
    buildModal(row);
  }

  // Abre modal vacío para crear
  function openNewModal() {
    creatingNew = true;

    const empty = {};
    currentHeaders.forEach(h => empty[h] = '');

    // Defaults de mes
    const monthName = new Intl.DateTimeFormat('es-CO', { month: 'long' }).format(new Date());
    const capital = (s)=> s.charAt(0).toUpperCase() + s.slice(1);
    if (currentHeaders.includes('Listado1') && !empty['Listado1']) empty['Listado1'] = capital(monthName);
    if (currentHeaders.includes('Listado')  && !empty['Listado'])  empty['Listado']  = capital(monthName);

    buildModal(empty, true);
  }

  // Construye el contenido del modal (se usa para editar y crear)
  function buildModal(row, isNew = false) {
    originalRow = isNew ? null : { ...row };

    const nombre = String(row['Nombre'] ?? '').trim() || (isNew ? 'Nuevo cliente' : 'Cliente');

    // Subtítulo base
    const canal = String(row['Canal de comunicación'] ?? '').trim();
    const grupo = String(row['Grupo'] ?? '').trim();
    const arte1 = String(row['Arte I'] ?? '').trim();
    const baseSubtitle = [grupo, arte1, canal].filter(Boolean).join(' · ');

    // 🟣 Mensajito de urgencia en el subtítulo
    const urgency = classifyUrgency(row);
    const urgencyMsg =
      urgency.type === 'urgent' ? ` · ⚡ Urgente${urgency.when ? ` (${urgency.when})` : ''}` :
      urgency.type === 'today'  ? ` · ⚡ Hoy` :
      urgency.type === 'soon'   ? ` · 📅 Agendado ${urgency.when}` :
      (urgency.type === 'ok' ? ' · ✅ Al día' : '');

    modalTitle.textContent = nombre;
    modalSubtitle.textContent = baseSubtitle + (urgencyMsg || '');

    const preferredOrder = [
      'ID','Listado1','Listado','Nombre','Celular/Teléfono','Correo Electrónico','Acudiente/Estudiante','Nombre de Estudiante',
      'Grupo','Edad',
      'Arte I','Instrumento/Estilo/Técnica I',
      'Arte II','Instrumento/Estilo/Técnica II',
      'Arte III','Instrumento/Estilo/Técnica III',
      'Modalidad','Curso/Plan','Ubicación','Fecha para contactar','Fecha y hora de contacto','Asesor','Comentario',
      '¿Tiene el instrumento?','Canal de comunicación','Prioridad'
    ];
    const headers = orderHeaders(currentHeaders, preferredOrder);

    modalBody.innerHTML = headers.map((h) => {
      const raw = row[h] ?? '';
      const id  = `f_${slug(h)}`;
      const headerName = String(h);
      const isLong = /comentario|ubicación|direccion|dirección/i.test(headerName);

      const arteIdx = ARTE_COLS.indexOf(headerName);
      const instIdx = INSTRUMENT_COLS.indexOf(headerName);

      let input;

      // 🔒 ID bloqueado siempre
      if (headerName === 'ID') {
        input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="text"
                 value="${escAttr(String(raw))}" readonly data-lock="1" />`;

      // Campo Acudiente/Estudiante
      } else if (headerName === 'Acudiente/Estudiante') {
        const opts = ["Acudiente","Estudiante"];
        input = selectHTML(id, h, opts, String(raw), !isNew);

      // Campo Edad
      } else if (headerName === 'Edad') {
        input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="number" min="0" value="${escAttr(String(raw))}" ${!isNew ? 'readonly' : ''} />`;

      // Campo ¿Tiene el instrumento?
      } else if (headerName === '¿Tiene el instrumento?') {
        const opts = ["Sí","No"];
        input = selectHTML(id, h, opts, String(raw), !isNew);

      // Selects fijos
      } else if (OPTIONS[headerName]) {
        const opts = OPTIONS[headerName];
        input = selectHTML(id, h, opts, String(raw), !isNew);

      // Arte I/II/III
      } else if (arteIdx !== -1) {
        input = selectHTML(id, h, OPTIONS.ARTE, String(raw), !isNew, { role:'arte', idx: arteIdx });

      // Instrumento dependiente
      } else if (instIdx !== -1) {
        const arteHeader = ARTE_COLS[instIdx];
        const currentArteVal = String(row[arteHeader] || '').trim();
        const opts = currentArteVal ? (ART_TO_OPTIONS[currentArteVal] || ART_TO_OPTIONS['Todos']) : [];
        input = selectHTML(id, h, opts, String(raw), !isNew, { role:'instrumento', idx: instIdx });

      // Fechas
      } else if (isDateHeader(headerName)) {
        const type = /hora/i.test(headerName) ? 'datetime-local' : 'date';
        const iso = toInputDateValue(String(raw), type);
        input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="${type}" value="${escAttr(iso)}" ${!isNew ? 'disabled' : ''} />`;

      // Textos largos
      } else if (isLong) {
        input = `<textarea id="${id}" name="${escAttr(h)}" class="in control" rows="3" ${!isNew ? 'readonly' : ''}>${esc(String(raw))}</textarea>`;

      // Resto
      } else {
        input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="text" value="${escAttr(String(raw))}" ${!isNew ? 'readonly' : ''} />`;
      }

      return `<div class="field">
        <div class="label">${esc(h)}</div>
        <div class="value">${input}</div>
      </div>`;
    }).join('');

    const footer = modal.querySelector('.modal-footer');
    footer.innerHTML = `
      <span id="saveStatus" class="muted"></span>
      <div style="margin-left:auto; display:flex; gap:.5rem;">
        ${isNew ? '' : '<button id="btnEdit" class="btn">Editar</button>'}
        <button id="btnSave" class="btn">${isNew ? 'Crear' : 'Guardar'}</button>
        <button class="btn" data-close="1">Cerrar</button>
      </div>`;

    const btnEdit    = document.getElementById('btnEdit');
    const btnSave    = document.getElementById('btnSave');
    const saveStatus = document.getElementById('saveStatus');

    // En modo edición (fila existente), primero hay que habilitar
    if (btnEdit) {
      btnSave.disabled = true;
      btnEdit.addEventListener('click', () => {
        modalBody.querySelectorAll('.control').forEach((el) => {
          if (el.dataset.lock === '1') return; // 👈 no tocar ID
          el.removeAttribute('readonly');
          el.removeAttribute('disabled');
        });
        btnSave.disabled = false;

        // Dependencias ArteX -> InstrumentoX
        wireArteDependencies();
      });
    } else {
      // En modo nuevo, los campos ya están habilitados
      wireArteDependencies();
    }

    // Guardar (crear o actualizar)
    btnSave.addEventListener('click', async () => {
      try {
        btnSave.disabled = true;
        saveStatus.textContent = creatingNew ? 'Creando…' : 'Guardando…';

        const formValues = {};
        modalBody.querySelectorAll('.control').forEach((el) => {
          const header = el.getAttribute('name');
          formValues[header] = el.value ?? '';
        });

        if (creatingNew) {
          // El ID lo genera tu flujo en Sheets si así lo decides
          const body = new URLSearchParams({
            mode: 'add',
            sheet: currentSheet,
            row: JSON.stringify(formValues)
          });
          const r = await fetch(API_BASE, { method:'POST', headers:{Accept:'application/json'}, body });
          let json = null, txt = '';
          try { json = await r.json(); } catch (_) { txt = await r.text(); }
          if (!r.ok || json?.ok !== true) {
            const msg = (json && json.error) || txt || `HTTP ${r.status}`;
            throw new Error(msg);
          }
          saveStatus.textContent = 'Creado ✓';
          await loadPage();
          closeModal();
          return;
        }

        // ----- EDICIÓN EXISTENTE -----
        const changes = diffObject(originalRow, formValues);
        delete changes['ID']; // nunca tocar ID

        if (!Object.keys(changes).length) {
          saveStatus.textContent = 'Sin cambios';
          btnSave.disabled = false;
          return;
        }
        if (!keyColumnInUse) throw new Error('No hay columna clave disponible en esta hoja.');
        const key = String(formValues[keyColumnInUse] ?? '').trim();
        if (!key) {
          saveStatus.textContent = `Falta valor en la columna clave: ${keyColumnInUse}`;
          btnSave.disabled = false;
          return;
        }

        const body = new URLSearchParams({
          mode: 'update',
          sheet: currentSheet,
          keyCol: keyColumnInUse,
          key,
          row: JSON.stringify(changes)
        });
        const r = await fetch(API_BASE, { method:'POST', headers:{Accept:'application/json'}, body });
        let json = null, txt = '';
        try { json = await r.json(); } catch (_) { txt = await r.text(); }
        if (!r.ok || json?.ok !== true) {
          const msg = (json && json.error) || txt || `HTTP ${r.status}`;
          throw new Error(msg);
        }
        saveStatus.textContent = 'Guardado ✓';
        await loadPage();
        closeModal();

      } catch (err) {
        console.error(err);
        saveStatus.textContent = `Error: ${err?.message || err}`;
        btnSave.disabled = false;
      }
    });

    modal.classList.remove('hidden');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
    modal.querySelector('.modal-close')?.focus();
  }

  function closeModal() {
    modal.classList.add('hidden');
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('modal-open');
    originalRow = null;
    creatingNew = false;
  }

  // Vincula cambios de ArteX -> repuebla InstrumentoX
  function wireArteDependencies(){
    modalBody.querySelectorAll('select[data-role="arte"]').forEach((arteSel) => {
      const idx = Number(arteSel.getAttribute('data-idx') || '0');
      arteSel.addEventListener('change', () => {
        const arte = arteSel.value || '';
        const newOpts = arte ? (ART_TO_OPTIONS[arte] || ART_TO_OPTIONS['Todos']) : [];
        const instSel = modalBody.querySelector(`select[data-role="instrumento"][data-idx="${idx}"]`);
        if (!instSel) return;
        const prev = instSel.value;
        instSel.innerHTML = selectOptionsHTML(newOpts, '');
        if (newOpts.includes(prev)) instSel.value = prev;
      });
    });
  }

  /* =============================
     6) Utilidades
     ============================= */
  function status(msg) { statusEl.textContent = msg; }

  function showError(err) {
    status(`⚠️ ${err?.message || err}`);
    console.error(err);
  }

  async function fetchJSON(url) {
    const r = await fetch(url, { headers: { Accept: 'application/json' } });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.json();
  }

  function esc(s) {
    return String(s ?? '').replace(/[&<>\"']/g, (c) =>
      ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c])
    );
  }
  function escAttr(s){ return String(s ?? '').replace(/"/g,'&quot;'); }

  function pretty(v) {
    if (v == null || v === '') return '';
    const s = String(v);
    if (/^música$/i.test(s)) return `<span class="badge music">${esc(s)}</span>`;
    if (/^(baile|danza)$/i.test(s)) return `<span class="badge dance">${esc(s)}</span>`;
    if (/^artes?\s?plásticas$/i.test(s)) return `<span class="badge arts">${esc(s)}</span>`;
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return `<a href="mailto:${escAttr(s)}">${esc(s)}</a>`;
    return esc(s);
  }

  function slug(h) {
    return String(h).toLowerCase()
      .normalize('NFD').replace(/\p{Diacritic}/gu,'')
      .replace(/[^a-z0-9]+/g,'-');
  }

  function diffObject(prev, next) {
    const out = {};
    const keys = new Set([...Object.keys(prev||{}), ...Object.keys(next||{})]);
    keys.forEach((k) => {
      const a = String(prev?.[k] ?? '');
      const b = String(next?.[k] ?? '');
      if (a !== b) out[k] = next[k];
    });
    return out;
  }

  function orderHeaders(all, preferred) {
    const seen = new Set();
    const first = preferred.filter((h) => all.includes(h) && !seen.has(h) && seen.add(h));
    const rest  = all.filter((h) => !seen.has(h));
    return [...first, ...rest];
  }

  function pickKeyColumn(headers) {
    return headers.includes('ID') ? 'ID' : '';
  }

  // --- Dates helpers ---
  function isDateHeader(h){
    return /fecha/i.test(h);
  }
  function toInputDateValue(raw, type){
    const s = String(raw || '').trim();
    if (!s) return '';
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T](\d{1,2}):(\d{2}))?$/);
    if (m){
      const dd = String(m[1]).padStart(2,'0');
      const mm = String(m[2]).padStart(2,'0');
      let yyyy = m[3].length === 2 ? ('20' + m[3]) : m[3];
      const date = `${yyyy}-${mm}-${dd}`;
      if (type === 'datetime-local' && m[4] && m[5]){
        const hh = String(m[4]).padStart(2,'0');
        const mi = String(m[5]).padStart(2,'0');
        return `${date}T${hh}:${mi}`;
      }
      return date;
    }
    if (/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2})?$/.test(s)) return s;
    // Fallback: Date.parse en español es frágil; se retorna tal cual si no matchea.
    return s;
  }

  // --- Select helpers con placeholder ---
  function selectHTML(id, header, options, current, disabled, dataAttrs){
    const attrs = [
      `id="${id}"`,
      `name="${escAttr(header)}"`,
      `class="in control"`,
      disabled ? 'disabled' : ''
    ];
    if (dataAttrs && dataAttrs.role) attrs.push(`data-role="${dataAttrs.role}"`);
    if (dataAttrs && dataAttrs.idx != null) attrs.push(`data-idx="${String(dataAttrs.idx)}"`);

    const html = [
      `<select ${attrs.join(' ')}>`,
      selectOptionsHTML(options, current),
      `</select>`
    ].join('');

    return html;
  }

  function selectOptionsHTML(options, current){
    const hasCurrent = current && options.includes(current);
    const optHtml = options.map(opt =>
      `<option value="${escAttr(opt)}" ${opt===current?'selected':''}>${esc(opt)}</option>`
    ).join('');
    return [
      `<option value="" ${(!current || !hasCurrent) ? 'selected' : ''}>— Selecciona —</option>`,
      optHtml
    ].join('');
  }

  function compareByKey(a, b, key) {
  const A = normalizeForSort(a?.[key]);
  const B = normalizeForSort(b?.[key]);

  // Nulls/vacíos al final
  const aNull = (A === null || A === undefined || A === '');
  const bNull = (B === null || B === undefined || B === '');
  if (aNull && bNull) return 0;
  if (aNull) return 1;
  if (bNull) return -1;

  // Fechas
  if (A instanceof Date && B instanceof Date) return A - B;

  // Números
  if (typeof A === 'number' && typeof B === 'number') return A - B;

  // Texto (insensible a mayúsculas/acentos)
  return String(A).localeCompare(String(B), 'es', { sensitivity: 'base', ignorePunctuation: true });
}

function normalizeForSort(v) {
  if (v == null) return '';
  const s = String(v).trim();

  // ¿Fecha tipo YYYY-MM-DD o DD/MM/YYYY? (aprovechamos tus helpers)
  const d = parseYMD(s);
  if (d instanceof Date && !isNaN(d)) return d;

  // ¿Número? (admite comas/puntos y símbolos simples)
  const numericish = s.replace(/[^\d.\-]/g, '');
  if (numericish && /^-?\d+(\.\d+)?$/.test(numericish)) {
    const n = Number(numericish);
    if (!isNaN(n)) return n;
  }

  return s.toLowerCase();
}

  /* =========================================================
   * 7) ADD-ON · Cola guiada "Contactar clientes"
   *     — Abre el MISMO modal estándar (edición) según el próximo pendiente.
   * ========================================================= */

  (function initContactButton() {
    try {
      const controlsBar = document.querySelector('.controls');
      if (!controlsBar || !btnNew) return;

      const btn = document.createElement('button');
      btn.id = 'btnContact';
      btn.type = 'button';
      btn.className = 'btn';
      btn.textContent = 'Contactar clientes';

      controlsBar.insertBefore(btn, btnNew);
      btn.addEventListener('click', openNextContactCard);
    } catch (e) {
      console.warn('No se pudo inyectar botón Contactar:', e);
    }
  })();

  // 🔁 Pide una tarjeta al backend y abre el MISMO modal de edición
  async function openNextContactCard() {
    try {
      const url = `${API_BASE}?mode=next&sheet=${encodeURIComponent(currentSheet)}&_ts=${Date.now()}`;
      const res = await fetchJSON(url);
      if (!res || !res.card || !res.card.ID) {
        status('No hay pendientes por ahora 🙌');
        alert('No hay pendientes por ahora 🙌');
        return;
      }

      // Buscar fila completa por ID (render actual o cache)
      const row = await getRowById(String(res.card.ID).trim());
      if (!row) {
        alert('No se encontró la fila en la hoja actual.');
        return;
      }

      // Abre modal estándar y activa edición para escribir de una
      openModal(row);
      setTimeout(() => {
        document.getElementById('btnEdit')?.click();
      }, 0);

    } catch (err) {
      console.error(err);
      status('Error pidiendo la siguiente tarjeta');
      alert('Error pidiendo la siguiente tarjeta');
    }
  }

  // Busca una fila por ID en las filas visibles o en el cache de toda la hoja
  async function getRowById(id) {
    const idKey = keyColumnInUse || 'ID';
    let row = displayedRows.find(r => String(r[idKey] ?? '').trim() === id);
    if (row) return row;

    let cache = allRowsCache.get(currentSheet);
    if (!cache) {
      cache = await loadAllRowsForSheet(currentSheet);
      allRowsCache.set(currentSheet, cache);
    }
    row = (cache.rows || []).find(r => String(r[idKey] ?? '').trim() === id);
    return row || null;
  }

  /* =========================================================
   * 8) Avisos (Urgentes y Agendados)
   *     - Calcula desde la hoja con cache.
   * ========================================================= */

  async function computeAndRenderAlerts() {
    if (!alertsEl) return;

    // Traer cache o cargar todo
    let cache = allRowsCache.get(currentSheet);
    if (!cache) {
      cache = await loadAllRowsForSheet(currentSheet);
      allRowsCache.set(currentSheet, cache);
    }

    const rows = cache.rows || [];

    const urgent = [];
    const soon   = [];

    rows.forEach(r => {
      const u = classifyUrgency(r);
      if (u.type === 'urgent' || u.type === 'today') urgent.push({row:r, when:u.whenSort || null});
      else if (u.type === 'soon') soon.push({row:r, when:u.whenSort || null});
    });

    // Orden: fecha más próxima primero; prioridad Alta antes
    const byWhen = (a,b) => {
      if (a.when && b.when && a.when.getTime() !== b.when.getTime()) return a.when < b.when ? -1 : 1;
      const ap = (String(a.row['Prioridad']||'').toLowerCase()==='alta')?0:1;
      const bp = (String(b.row['Prioridad']||'').toLowerCase()==='alta')?0:1;
      return ap - bp;
    };
    urgent.sort(byWhen);
    soon.sort(byWhen);

    renderAlertsUI(
      urgent.slice(0,5).map(x=>x.row),
      soon.slice(0,5).map(x=>x.row),
      urgent.length, soon.length
    );
  }

  function renderAlertsUI(urgentTop, soonTop, uCount, sCount){
    if (!alertsEl) return;

    if (uCount === 0 && sCount === 0) {
      alertsEl.classList.add('hidden');
      alertsEl.innerHTML = '';
      return;
    }

    const makeItem = (r) => {
      const name = (String(r['Nombre']||'').trim()) || '(Sin nombre)';
      const prio = (String(r['Prioridad']||'').trim()) || '';
      const canal = (String(r['Canal de comunicación'] || r['Canal'] || '').trim());
      const when = minDate(parseYMD(r['Fecha para contactar']), parseYMD(r['Siguiente Contacto (calc)']));
      const whenTxt = when ? formatYMD(when) : '—';
      const id = String(r['ID']||'').trim();

      return `
        <div class="alert-item">
          <strong>${esc(name)}</strong>
          <span class="meta">· ${esc(prio || '—')}${canal?` · ${esc(canal)}`:''} · ${esc(whenTxt)}</span>
          <button class="btn" data-openid="${escAttr(id)}">Abrir</button>
        </div>
      `;
    };

    alertsEl.innerHTML = `
      <div class="alert-card">
        <div class="alert-title">
          <span>📞 Contactar clientes</span>
          <span class="kebab">·</span>
          <div class="pills">
            <span class="pill urgent">Urgentes: ${uCount}</span>
            <span class="pill soon">Agendados (≤3 días): ${sCount}</span>
          </div>
        </div>
        <div class="alert-list">
          ${urgentTop.map(makeItem).join('')}
          ${soonTop.map(makeItem).join('')}
        </div>
      </div>
    `;
    alertsEl.classList.remove('hidden');

    // Abrir modal desde el aviso
    alertsEl.querySelectorAll('[data-openid]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-openid') || '';
        const row = await getRowById(id);
        if (row) {
          openModal(row);
          setTimeout(()=>document.getElementById('btnEdit')?.click(), 0);
        } else {
          alert('No se encontró la fila seleccionada.');
        }
      });
    });
  }

  // -------- Urgencia: lógica común para modal y avisos --------
  function classifyUrgency(row){
    const prio = (String(row['Prioridad'] ?? '')).trim().toLowerCase();
    const f1   = parseYMD(String(row['Fecha para contactar'] ?? ''));
    const f2   = parseYMD(String(row['Siguiente Contacto (calc)'] ?? ''));
    const when = minDate(f1, f2);

    const today = toYMD(new Date());
    const in3   = addDays(today, 3);

    const dueOrPast = (d) => d && d <= today;
    const within3   = (d) => d && d > today && d <= in3;

    if (prio === 'alta' && dueOrPast(when)) return { type:'urgent', when: formatIf(when), whenSort: when };
    if (prio === 'alta' && !when)           return { type:'urgent', when: '', whenSort: today };
    if (dueOrPast(when))                    return { type: when && when.getTime() === today.getTime() ? 'today' : 'urgent', when: formatIf(when), whenSort: when || today };
    if (within3(when))                      return { type:'soon', when: formatIf(when), whenSort: when };
    return { type:'ok', when:'', whenSort: null };
  }

  /* =============================
     9) Helpers de fecha reutilizados
     ============================= */
  function toYMD(d){ return d instanceof Date ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null; }
  function addDays(d, n){ const x = new Date(d); x.setDate(x.getDate()+n); return toYMD(x); }
  function minDate(a,b){ if (a && b) return a<b?a:b; return a||b||null; }
  function parseYMD(s){
    if (!s) return null;
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return toYMD(new Date(Number(iso[1]), Number(iso[2])-1, Number(iso[3])));
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (dmy){
      const dd = Number(dmy[1]), mm = Number(dmy[2])-1, yyyy = dmy[3].length===2 ? Number('20'+dmy[3]) : Number(dmy[3]);
      return toYMD(new Date(yyyy, mm, dd));
    }
    return null;
  }
  function pad2(n){ return String(n).padStart(2,'0'); }
  function formatYMD(d){ return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; }
  function formatIf(d){ return d ? formatYMD(d) : ''; }

})();
