/* =============================
   app.js — Seguimiento Comercial Musicala
   Paginación + búsqueda global + modal editable (v1.2)
   + Botón WhatsApp debajo del celular (robusto)
   + Fix click con ordenamiento (rowindex consistente)
   ============================= */

'use strict';

(async function () {
  // ------- DOM -------
  const statusEl   = document.getElementById('status');
  const table      = document.getElementById('dataTable');
  const sheetSel   = document.getElementById('sheetSelect');
  const searchBox  = document.getElementById('search');
  const btnNew     = document.getElementById('btnNew');
  const btnSyncFirebase = document.getElementById('btnSyncFirebase');
  const btnSyncSheets = document.getElementById('btnSyncSheets');
  const btnGoogleLogin = document.getElementById('btnGoogleLogin');
  const firebaseStatusEl = document.getElementById('firebaseStatus');
  const progressPanel = document.getElementById('progressPanel');
  const progressTitle = document.getElementById('progressTitle');
  const progressPercent = document.getElementById('progressPercent');
  const progressBar = document.getElementById('progressBar');
  const progressText = document.getElementById('progressText');

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
  const ALLOWED_EMAILS = [
    'alekcaballeromusic@gmail.com',
    'catalina.medina.leal@gmail.com',
    'musicalaasesor@gmail.com',
    'imusicala@gmail.com'
  ];
  const firebaseDb = initFirebaseCache();
  const authReady = waitForFirebaseAuth();

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
    Asesor: ["Alek Caballero","Catalina Medina","Camila Rodríguez","Liceth Rincón"],
    "Canal de comunicación": ["Llamada","WhatsApp","Keybe","Instagram","Facebook","Wix","TikTok"],
    Prioridad: ["Alta","Media","Baja"]
  };

  // ====== Dependencias Arte/Instrumento ======
  const ART_TO_OPTIONS = {
    "Música": [
      "Piano","Guitarra","Canto","Violín","Batería","Cello","Bajo eléctrico","Ukelele",
      "Acordeón","Bandola","Iniciación Musical","Teoría","Composición música",
      "Flauta Dulce","Flauta Traversa","Saxofón","Requinto","Trombón","Trompeta","Percusión","Xilófono","Jazz"
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

  // OJO: displayedRows SIEMPRE debe ser exactamente lo que está renderizado
  let displayedRows  = []; // filas actualmente renderizadas (página o filtradas y/o ordenadas)
  btnGoogleLogin?.addEventListener('click', () => signInWithGoogle());

  /* =============================
     1) Cargar hojas + primera página
     ============================= */
  lockDataUI(true);
  const authUser = await authReady;
  if (!authUser || !isAllowedUser(authUser)) {
    status(authUser ? 'Esta cuenta no tiene permiso para cargar la base de datos.' : 'Ingresa con Google para cargar la base de datos.');
    renderLockedTable();
    return;
  }
  lockDataUI(false);

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

  await runDailySyncIfNeeded();
  await loadPage();

  /* =============================
     2) Eventos UI
     ============================= */
  sheetSel.addEventListener('change', async (e) => {
    currentSheet = e.target.value;
    offset = 0;
    searchBox.value = '';
    searchActive = false;
    sortState = { key: null, dir: 1 };
    pagerEl?.classList.remove('hidden');
    await loadPage();
  });

  btnPrev?.addEventListener('click', async () => {
    if (searchActive) return;
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
  btnSyncFirebase?.addEventListener('click', async () => {
    try {
      btnSyncFirebase.disabled = true;
      await syncCurrentSheetToFirebase(true);
    } finally {
      btnSyncFirebase.disabled = false;
    }
  });
  btnSyncSheets?.addEventListener('click', async () => {
    try {
      btnSyncSheets.disabled = true;
      await syncPendingFirebaseToSheets(true);
    } finally {
      btnSyncSheets.disabled = false;
    }
  });
  // Delegación: click en el "Nombre" abre modal de edición
  table.addEventListener('click', (e) => {
    const btn = e.target.closest('.name-link');
    if (!btn) return;
    const idx = Number(btn.dataset.rowindex);
    const row = displayedRows[idx]; // ✅ ahora coincide con lo renderizado
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
    const cached = await loadSheetFromFirebase(currentSheet);
    if (cached) {
      currentHeaders = cached.headers;
      total = cached.total;
      currentRows = (cached.rows || []).slice(offset, offset + limit);
      keyColumnInUse = pickKeyColumn(currentHeaders);
      renderTable(currentHeaders, currentRows);

      const page  = Math.floor(offset / limit) + 1;
      const pages = Math.max(1, Math.ceil(total / limit));
      if (pageInfo) pageInfo.textContent = `Página ${page.toLocaleString()} de ${pages.toLocaleString()} — ${total.toLocaleString()} registros`;
      if (btnPrev) btnPrev.disabled = offset <= 0;
      if (btnNext) btnNext.disabled = offset + limit >= total;
      status(`Listo. Mostrando ${currentRows.length.toLocaleString()} de ${total.toLocaleString()} registros.`);
      allRowsCache.set(currentSheet, cached);
      await computeAndRenderAlerts();
      return;
    }

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

    // ✅ Render y seteo de displayedRows dentro de renderTable (incluye ordenamiento)
    renderTable(currentHeaders, currentRows);

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

    // Si hay una columna activa, ordena una copia
    let toRender = rows;
    if (sortState.key) {
      const key = sortState.key;
      const dir = sortState.dir;

      toRender = [...rows].map((r, i) => ({ r, i })) // estable
        .sort((a, b) => {
          const cmp = compareByKey(a.r, b.r, key);
          return (cmp !== 0) ? (cmp * dir) : (a.i - b.i);
        })
        .map(x => x.r);
    }

    // ✅ LO MÁS IMPORTANTE: lo que se ve = displayedRows
    displayedRows = toRender;

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

    // ancho mínimo para scroll
    const pxPerCol = 140;
    table.style.minWidth = (headers.length * pxPerCol) + 'px';

    // listeners ordenar
    table.querySelectorAll('thead th.th-sortable').forEach(th => {
      th.addEventListener('click', () => {
        const key = th.getAttribute('data-key') || '';
        if (!key) return;
        if (sortState.key === key) sortState.dir = sortState.dir === 1 ? -1 : 1;
        else { sortState.key = key; sortState.dir = 1; }
        renderTable(headers, rows);
      });
    });
  }

  // Carga TODA la hoja y la junta
  async function loadAllRowsForSheet(sheetName) {
    const cached = await loadSheetFromFirebase(sheetName);
    if (cached) return cached;

    return fetchAllRowsFromSheets(sheetName, true);
  }

  async function fetchAllRowsFromSheets(sheetName, mirrorToFirebase = false) {
    const first = await fetchJSON(`${API_BASE}?mode=data&sheet=${encodeURIComponent(sheetName)}&limit=1&offset=0&_ts=${Date.now()}`);
    const headers = first.headers || [];

    const pageLimit = 1000;
    let off = 0;
    let tot = 0;
    let all = [];

    while (true) {
      const url = `${API_BASE}?mode=data&sheet=${encodeURIComponent(sheetName)}&limit=${pageLimit}&offset=${off}&_ts=${Date.now()}`;
      const res = await fetchJSON(url);
      const rows = res.rows || [];
      tot = parseInt(res.total || rows.length, 10);
      all = all.concat(rows);
      if (mirrorToFirebase && tot > 0) {
        const pct = 55 + Math.min(25, Math.round((all.length / tot) * 25));
        setProgress(pct, `Leyendo registros (${Math.min(all.length, tot).toLocaleString()} de ${tot.toLocaleString()})...`);
      }
      off += pageLimit;
      if (all.length >= tot || rows.length === 0) break;
    }

    const payload = { headers, rows: all, total: tot, ts: Date.now(), source: 'sheets' };
    if (mirrorToFirebase) {
      setProgress(82, 'Guardando datos para una carga más rápida...');
      await saveSheetToFirebase(sheetName, payload).catch((err) => console.warn('No se pudo actualizar Firebase:', err));
    }
    return payload;
  }

  /* =============================
     4) Búsqueda GLOBAL
     ============================= */
  async function applySearch() {
    const q = (searchBox?.value || '').trim();
    if (!q) {
      searchActive = false;
      pagerEl?.classList.remove('hidden');
      status(`Listo. Clave: “${keyColumnInUse || '—'}”. Mostrando ${currentRows.length.toLocaleString()} de ${total.toLocaleString()} registros.`);
      renderTable(currentHeaders, currentRows);

      const page  = Math.floor(offset / limit) + 1;
      const pages = Math.max(1, Math.ceil(total / limit));
      if (pageInfo) pageInfo.textContent = `Página ${page.toLocaleString()} de ${pages.toLocaleString()} — ${total.toLocaleString()} registros.`;
      return;
    }

    try {
      searchActive = true;
      pagerEl?.classList.add('hidden');
      status('Buscando en toda la hoja…');

      let cache = allRowsCache.get(currentSheet);
      if (!cache) {
        cache = await loadAllRowsForSheet(currentSheet);
        allRowsCache.set(currentSheet, cache);
      }

      const norm = (s) => String(s ?? '')
        .toLowerCase()
        .normalize('NFD').replace(/\p{Diacritic}/gu,'')
        .trim();

      const terms = norm(q).split(/\s+/).filter(Boolean);

      const rows = cache.rows || [];
      const headers = cache.headers || currentHeaders;

      const filtered = rows.filter(r => {
        const blob = norm(headers.map(h => r[h]).join(' | '));
        return terms.every(t => blob.includes(t));
      });

      status(`Resultado: ${filtered.length.toLocaleString()} coincidencia(s) en “${currentSheet}”.`);
      if (pageInfo) pageInfo.textContent = `Filtrado — ${filtered.length.toLocaleString()} coincidencia(s)`;

      renderTable(headers, filtered);

    } catch (err) {
      console.error(err);
      status(`⚠️ Error en la búsqueda: ${err?.message || err}`);
    }
  }

  /* =============================
     5) Modales
     ============================= */
  function openModal(row) {
    creatingNew = false;
    buildModal(row);
  }

  function openNewModal() {
    creatingNew = true;

    const empty = {};
    currentHeaders.forEach(h => empty[h] = '');

    const monthName = new Intl.DateTimeFormat('es-CO', { month: 'long' }).format(new Date());
    const capital = (s)=> s.charAt(0).toUpperCase() + s.slice(1);
    if (currentHeaders.includes('Listado1') && !empty['Listado1']) empty['Listado1'] = capital(monthName);
    if (currentHeaders.includes('Listado')  && !empty['Listado'])  empty['Listado']  = capital(monthName);

    buildModal(empty, true);
  }

  function buildModal(row, isNew = false) {
    originalRow = isNew ? null : { ...row };

    const nombre = String(row['Nombre'] ?? '').trim() || (isNew ? 'Nuevo cliente' : 'Cliente');

    const canal = String(row['Canal de comunicación'] ?? '').trim();
    const grupo = String(row['Grupo'] ?? '').trim();
    const arte1 = String(row['Arte I'] ?? '').trim();
    const baseSubtitle = [grupo, arte1, canal].filter(Boolean).join(' · ');

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
      'Modalidad','Curso/Plan','Ubicación','Asesor','Canal de comunicación','Fecha y hora de contacto','Fecha para contactar','Comentario',
      '¿Tiene el instrumento?','Canal de comunicación','Prioridad'
    ];
    const headers = orderHeaders(currentHeaders, preferredOrder);

    // Render fields (✅ aquí NO se consulta el DOM todavía)
    modalBody.innerHTML = headers.map((h) => renderFieldHTML(h, row[h] ?? '', row, isNew)).join('');

    // ✅ Ahora sí: como ya existe el DOM del modal, conectamos WhatsApp + dependencias
    wireWhatsAppButton();
    if (isNew) wireArteDependencies(); // en nuevo, ya editable

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

    if (btnEdit) {
      btnSave.disabled = true;
      btnEdit.addEventListener('click', () => {
        modalBody.querySelectorAll('.control').forEach((el) => {
          if (el.dataset.lock === '1') return;
          el.removeAttribute('readonly');
          el.removeAttribute('disabled');
        });
        btnSave.disabled = false;
        wireArteDependencies();
        wireWhatsAppButton(); // por si se habilita el input
      });
    } else {
      wireArteDependencies();
    }

    btnSave.addEventListener('click', async () => {
      try {
        btnSave.disabled = true;
        saveStatus.textContent = creatingNew ? 'Creando…' : 'Guardando…';

        const formValues = {};
        modalBody.querySelectorAll('.control').forEach((el) => {
          const header = el.getAttribute('name');
          formValues[header] = el.value ?? '';
        });
        validateTrackingDates(formValues);

        if (creatingNew) {
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
          allRowsCache.delete(currentSheet);
          await syncCurrentSheetToFirebase(false);
          await loadPage();
          closeModal();
          return;
        }

        const changes = diffObject(originalRow, formValues);
        delete changes['ID'];

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

        await saveRowToFirebase(currentSheet, { ...originalRow, ...formValues }, true);
        saveStatus.textContent = 'Guardado ✓ Pendiente de enviar a principal';
        allRowsCache.delete(currentSheet);
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

  // Render de un campo
  function renderFieldHTML(h, raw, fullRow, isNew) {
    const id  = `f_${slug(h)}`;
    const headerName = String(h);
    const labelText = getFieldLabel(headerName);
    const isLong = /comentario|ubicación|direccion|dirección/i.test(headerName);

    const arteIdx = ARTE_COLS.indexOf(headerName);
    const instIdx = INSTRUMENT_COLS.indexOf(headerName);

    const fieldClass = isTrackingDateField(headerName) ? ' field-tracking-date' : '';

    let input = '';

    // 🔒 ID bloqueado siempre
    if (headerName === 'ID') {
      input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="text"
               value="${escAttr(String(raw))}" readonly data-lock="1" />`;

    } else if (headerName === 'Acudiente/Estudiante') {
      input = selectHTML(id, h, ["Acudiente","Estudiante"], String(raw), !isNew);

    } else if (headerName === 'Edad') {
      input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="number" min="0"
               value="${escAttr(String(raw))}" ${!isNew ? 'readonly' : ''} />`;

    } else if (headerName === '¿Tiene el instrumento?') {
      input = selectHTML(id, h, ["Sí","No"], String(raw), !isNew);

    } else if (OPTIONS[headerName]) {
      input = selectHTML(id, h, OPTIONS[headerName], String(raw), !isNew);

    } else if (arteIdx !== -1) {
      input = selectHTML(id, h, OPTIONS.ARTE, String(raw), !isNew, { role:'arte', idx: arteIdx });

    } else if (instIdx !== -1) {
      const arteHeader = ARTE_COLS[instIdx];
      const currentArteVal = String(fullRow[arteHeader] || '').trim();
      const opts = currentArteVal ? (ART_TO_OPTIONS[currentArteVal] || ART_TO_OPTIONS['Todos']) : [];
      input = selectHTML(id, h, opts, String(raw), !isNew, { role:'instrumento', idx: instIdx });

    } else if (isDateHeader(headerName)) {
      const type = /hora/i.test(headerName) ? 'datetime-local' : 'date';
      const autoRaw = isContactDateTimeField(headerName)
        ? currentDateTimeLocal()
        : String(raw);
      const iso = toInputDateValue(autoRaw, type);
      const minAttr = headerName === 'Fecha para contactar' ? ` min="${todayISODate()}"` : '';
      input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="${type}"
               value="${escAttr(iso)}"${minAttr} ${!isNew ? 'disabled' : ''} />`;

    } else if (isLong) {
      input = `<textarea id="${id}" name="${escAttr(h)}" class="in control" rows="3"
               ${!isNew ? 'readonly' : ''}>${esc(String(raw))}</textarea>`;

    } else {
      input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="text"
               value="${escAttr(String(raw))}" ${!isNew ? 'readonly' : ''} />`;
    }

    // ✅ Botón WhatsApp: se renderiza aquí mismo (siempre queda debajo)
    const isPhone = headerName.trim() === 'Celular/Teléfono';
    const waHtml = isPhone
      ? `<a class="btn btn-whatsapp" id="waBtn" target="_blank" rel="noopener" style="margin-top:.5rem; display:none;">💬 WhatsApp</a>`
      : '';

    return `<div class="field${fieldClass}">
      <div class="label">${esc(labelText)}</div>
      <div class="value">${input}${waHtml}</div>
    </div>`;
  }

  function getFieldLabel(headerName) {
    if (headerName === 'Fecha para contactar') return 'Fecha para contactar nuevamente';
    if (isContactDateTimeField(headerName)) return 'Fecha y hora en que se contactó';
    return headerName;
  }

  function isTrackingDateField(headerName) {
    return headerName === 'Fecha para contactar' || isContactDateTimeField(headerName);
  }

  function isContactDateTimeField(headerName) {
    const normalized = String(headerName || '')
      .normalize('NFD').replace(/\p{Diacritic}/gu, '')
      .trim()
      .toLowerCase();
    return normalized === 'fecha y hora de contacto';
  }

  function validateTrackingDates(formValues) {
    const followUp = String(formValues['Fecha para contactar'] || '').trim();
    if (!followUp) return;

    const followUpDate = toInputDateValue(followUp, 'date').slice(0, 10);
    if (/^\d{4}-\d{2}-\d{2}$/.test(followUpDate) && followUpDate < todayISODate()) {
      const field = modalBody.querySelector('[name="Fecha para contactar"]');
      field?.focus();
      throw new Error('Hey, no: la fecha para contactar nuevamente está en el pasado. Debe ser hoy o una fecha futura.');
    }
  }

  function closeModal() {
    modal.classList.add('hidden');
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('modal-open');
    originalRow = null;
    creatingNew = false;
  }

  // ✅ WhatsApp wiring (ya con DOM listo)
  function wireWhatsAppButton() {
    const phoneEl = modalBody.querySelector('[name="Celular/Teléfono"]');
    const waBtn   = modalBody.querySelector('#waBtn');
    if (!phoneEl || !waBtn) return;

    const update = () => {
      const raw = phoneEl.value || '';
      const clean = raw.replace(/\D/g, '');
      if (clean.length >= 7) {
        const num = clean.startsWith('57') ? clean.slice(2) : clean;
        waBtn.href = `https://wa.me/57${num}`;
        waBtn.style.display = 'inline-flex';
      } else {
        waBtn.style.display = 'none';
      }
    };

    // evitar duplicar listeners (por si abres/cierra modal mil veces)
    if (!phoneEl.dataset.waBound) {
      phoneEl.addEventListener('input', update);
      phoneEl.dataset.waBound = '1';
    }
    update();
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
     6) Firebase como cache de lectura
     ============================= */
  function initFirebaseCache() {
    const cfg = window.FIREBASE_CONFIG || {};
    const ready = cfg.apiKey && cfg.authDomain && cfg.projectId && window.firebase?.initializeApp;
    if (!ready) {
      setFirebaseStatus('No se pudo preparar el acceso seguro.');
      btnSyncFirebase?.setAttribute('disabled', 'disabled');
      return null;
    }

    try {
      const app = window.firebase.apps?.length ? window.firebase.app() : window.firebase.initializeApp(cfg);
      wireFirebaseAuth();
      setFirebaseStatus('Ingresa con Google para ver la base de datos.');
      return window.firebase.firestore(app);
    } catch (err) {
      console.error(err);
      setFirebaseStatus('No se pudo iniciar el acceso seguro.');
      btnSyncFirebase?.setAttribute('disabled', 'disabled');
      return null;
    }
  }

  function wireFirebaseAuth() {
    if (!window.firebase?.auth) return;
    const auth = window.firebase.auth();
    btnSyncFirebase?.setAttribute('disabled', 'disabled');

    auth.onAuthStateChanged((user) => {
      if (user?.email && isAllowedUser(user)) {
        btnSyncFirebase?.removeAttribute('disabled');
        btnSyncSheets?.removeAttribute('disabled');
        if (btnGoogleLogin) btnGoogleLogin.textContent = user.email;
        setFirebaseStatus(`Sesión iniciada: ${user.email}. Ya puedes consultar y actualizar datos.`);
      } else if (user?.email) {
        btnSyncFirebase?.setAttribute('disabled', 'disabled');
        btnSyncSheets?.setAttribute('disabled', 'disabled');
        if (btnGoogleLogin) btnGoogleLogin.textContent = user.email;
        setFirebaseStatus(`La cuenta ${user.email} no esta autorizada para esta base.`);
      } else {
        btnSyncFirebase?.setAttribute('disabled', 'disabled');
        btnSyncSheets?.setAttribute('disabled', 'disabled');
        if (btnGoogleLogin) btnGoogleLogin.textContent = 'Ingresar con Google';
        setFirebaseStatus('Ingresa con Google para ver la base de datos.');
      }
    });
  }

  function isAllowedUser(user) {
    const email = String(user?.email || '').toLowerCase();
    return ALLOWED_EMAILS.includes(email);
  }

  function waitForFirebaseAuth() {
    if (!window.firebase?.auth) {
      setFirebaseStatus('No se pudo preparar el inicio de sesión.');
      return Promise.resolve(null);
    }

    return new Promise((resolve) => {
      const stop = window.firebase.auth().onAuthStateChanged((user) => {
        stop();
        resolve(user || null);
      });
    });
  }

  async function signInWithGoogle() {
    if (!window.firebase?.auth) {
      setFirebaseStatus('No se pudo abrir el inicio de sesión.');
      return;
    }
    try {
      const provider = new window.firebase.auth.GoogleAuthProvider();
      await window.firebase.auth().signInWithPopup(provider);
      window.location.reload();
    } catch (err) {
      console.error(err);
      setFirebaseStatus('No se pudo iniciar sesión. Revisa que uses una cuenta autorizada.');
    }
  }

  function lockDataUI(locked) {
    [sheetSel, searchBox, btnNew, btnSyncFirebase, btnSyncSheets].forEach((el) => {
      if (!el) return;
      el.disabled = locked || ((el === btnSyncFirebase || el === btnSyncSheets) && !firebaseDb);
    });
    if (locked) {
      sheetSel.innerHTML = '<option>Inicia sesion</option>';
      searchBox.value = '';
      alertsEl?.classList.add('hidden');
      pagerEl?.classList.add('hidden');
    } else {
      pagerEl?.classList.remove('hidden');
    }
  }

  function renderLockedTable() {
    table.innerHTML = `
      <tbody>
        <tr>
          <td class="locked-cell">
            Ingresa con Google para ver la base de datos.
          </td>
        </tr>
      </tbody>
    `;
    table.style.minWidth = '100%';
  }

  function setFirebaseStatus(msg) {
    if (firebaseStatusEl) firebaseStatusEl.textContent = msg;
  }

  function showProgress(title, text, percent = 0) {
    if (!progressPanel) return;
    progressPanel.classList.remove('hidden');
    if (progressTitle) progressTitle.textContent = title;
    setProgress(percent, text);
  }

  function setProgress(percent, text) {
    const safe = Math.max(0, Math.min(100, Number(percent) || 0));
    if (progressPercent) progressPercent.textContent = `${safe}%`;
    if (progressBar) progressBar.style.width = `${safe}%`;
    if (progressText && text) progressText.textContent = text;
  }

  function hideProgress() {
    progressPanel?.classList.add('hidden');
  }

  function sheetDoc(sheetName) {
    if (!firebaseDb) return null;
    return firebaseDb.collection('sheetCache').doc(slug(sheetName) || 'hoja');
  }

  function rowDocId(row, index) {
    const id = String(row?.ID ?? '').trim();
    return id ? slug(id) : `fila-${index + 1}`;
  }

  async function loadSheetFromFirebase(sheetName) {
    if (!firebaseDb) return null;

    try {
      const doc = sheetDoc(sheetName);
      const metaSnap = await doc.get();
      if (!metaSnap.exists) return null;

      const rowSnap = await doc.collection('rows').get();
      if (rowSnap.empty) return null;

      const meta = metaSnap.data() || {};
      const rows = rowSnap.docs.map(d => d.data().data || d.data());
      const headers = Array.isArray(meta.headers) && meta.headers.length
        ? meta.headers
        : Array.from(rows.reduce((set, row) => {
            Object.keys(row || {}).forEach(k => set.add(k));
            return set;
          }, new Set()));

      setFirebaseStatus(`Base lista: ${rows.length.toLocaleString()} registros disponibles.`);
      return { headers, rows, total: rows.length, ts: meta.updatedAt || Date.now(), source: 'firebase' };
    } catch (err) {
      console.warn('No se pudo leer la base rapida, se usara respaldo:', err);
      setFirebaseStatus('Cargando información actualizada...');
      return null;
    }
  }

  async function saveSheetToFirebase(sheetName, payload) {
    if (!firebaseDb || !payload?.rows?.length) return;

    const doc = sheetDoc(sheetName);
    const batchSize = 400;
    const rows = payload.rows || [];

    await doc.set({
      sheetName,
      headers: payload.headers || [],
      total: rows.length,
      updatedAt: Date.now(),
      source: 'sheets'
    }, { merge: true });

    for (let start = 0; start < rows.length; start += batchSize) {
      const batch = firebaseDb.batch();
      rows.slice(start, start + batchSize).forEach((row, i) => {
        batch.set(doc.collection('rows').doc(rowDocId(row, start + i)), {
          data: row,
          updatedAt: Date.now(),
          pendingSheetSync: false,
          pendingReason: '',
          syncedAt: Date.now()
        }, { merge: true });
      });
      await batch.commit();
    }

    setFirebaseStatus(`Datos actualizados: ${rows.length.toLocaleString()} registros disponibles.`);
  }

  async function saveRowToFirebase(sheetName, row, pendingSheetSync = false) {
    if (!firebaseDb) throw new Error('No se pudo guardar el cambio rapido.');
    const doc = sheetDoc(sheetName);
    const id = rowDocId(row, 0);
    if (!id) throw new Error('El registro no tiene ID.');

    await doc.collection('rows').doc(id).set({
      data: row,
      updatedAt: Date.now(),
      pendingSheetSync,
      pendingReason: pendingSheetSync ? 'edit' : '',
      syncedAt: pendingSheetSync ? null : Date.now()
    }, { merge: true });

    await doc.set({
      sheetName,
      updatedAt: Date.now()
    }, { merge: true });
  }

  async function updateRowInSheets(sheetName, row) {
    const keyCol = keyColumnInUse || pickKeyColumn(Object.keys(row || {}));
    const key = String(row?.[keyCol] ?? '').trim();
    if (!keyCol || !key) throw new Error('No se encontro ID para actualizar la base principal.');

    const changes = { ...row };
    delete changes.ID;

    const body = new URLSearchParams({
      mode: 'update',
      sheet: sheetName,
      keyCol,
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
  }

  async function syncPendingFirebaseToSheets(showMessages = false) {
    if (!firebaseDb) {
      setFirebaseStatus('No se pudo guardar en la base principal en este momento.');
      return;
    }

    const doc = sheetDoc(currentSheet);
    const snap = await doc.collection('rows').where('pendingSheetSync', '==', true).get();
    if (snap.empty) {
      if (showMessages) status('No hay cambios pendientes por guardar en la base principal.');
      return;
    }

    if (showMessages) status(`Guardando ${snap.size.toLocaleString()} cambio(s) en la base principal...`);

    let done = 0;
    for (const rowDoc of snap.docs) {
      const payload = rowDoc.data() || {};
      const row = payload.data || {};
      setProgress(Math.round((done / snap.size) * 100), `Guardando cambios en la base principal (${done + 1} de ${snap.size})...`);
      await updateRowInSheets(currentSheet, row);
      await rowDoc.ref.set({
        pendingSheetSync: false,
        pendingReason: '',
        syncedAt: Date.now()
      }, { merge: true });
      done += 1;
    }

    allRowsCache.delete(currentSheet);
    if (showMessages) status(`${done.toLocaleString()} cambio(s) guardados en la base principal.`);
  }

  async function runDailySyncIfNeeded() {
    if (!firebaseDb || !currentSheet) return;

    const doc = sheetDoc(currentSheet);
    const today = formatYMD(toYMD(new Date()));
    const snap = await doc.get().catch(() => null);
    const meta = snap?.exists ? (snap.data() || {}) : {};
    if (meta.dailySyncDate === today) return;

    showProgress('Actualizando datos diarios', 'Preparando actualización segura...', 8);
    status('Espera un momento, actualizando los datos del día...');
    setFirebaseStatus('Actualizando datos diarios. Esto puede tardar un momento.');

    await doc.set({
      sheetName: currentSheet,
      dailySyncStartedAt: Date.now()
    }, { merge: true });

    setProgress(20, 'Guardando cambios pendientes en la base principal...');
    await syncPendingFirebaseToSheets(false);
    setProgress(55, 'Trayendo la información más reciente...');
    await syncCurrentSheetToFirebase(false);

    await doc.set({
      sheetName: currentSheet,
      dailySyncDate: today,
      dailySyncFinishedAt: Date.now()
    }, { merge: true });

    setProgress(100, 'Listo. Cargando la base...');
    status('Datos diarios actualizados. Cargando base...');
    setFirebaseStatus('Datos del día actualizados correctamente.');
    window.setTimeout(hideProgress, 900);
  }

  async function syncCurrentSheetToFirebase(showMessages = false) {
    if (!firebaseDb) {
      setFirebaseStatus('No se pudo actualizar la base de datos en este momento.');
      return;
    }

    if (showMessages) {
      showProgress('Actualizando datos', 'Trayendo información actualizada...', 10);
      status(`Actualizando datos de "${currentSheet}"...`);
    }
    const fresh = await fetchAllRowsFromSheets(currentSheet, true);
    allRowsCache.set(currentSheet, { ...fresh, source: 'firebase' });
    if (showMessages) setProgress(85, 'Recalculando avisos de seguimiento...');
    await computeAndRenderAlerts();
    if (showMessages) {
      setProgress(100, 'Datos actualizados correctamente.');
      status(`Datos actualizados para "${currentSheet}".`);
      window.setTimeout(hideProgress, 900);
    }
  }

  /* =============================
     7) Utilidades
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
      ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])
    );
  }
  function escAttr(s){ return String(s ?? '').replace(/"/g,'&quot;'); }

  function pretty(v) {
    if (v == null || v === '') return '';
    const s = String(v);
    if (/^música$/i.test(s)) return `<span class="badge music">${esc(s)}</span>`;
    if (/^(baile|danza)$/i.test(s)) return `<span class="badge dance">${esc(s)}</span>`;
    if (/^artes?\s?plásticas$/i.test(s)) return `<span class="badge arts">${esc(s)}</span>`;
    if (/^teatro$/i.test(s)) return `<span class="badge theatre">${esc(s)}</span>`;
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
    const first = preferred.filter((h) => all.includes(h) && !seen.has(h) && (seen.add(h), true));
    const rest  = all.filter((h) => !seen.has(h));
    return [...first, ...rest];
  }

  function pickKeyColumn(headers) {
    // Solo ID
    return headers.includes('ID') ? 'ID' : '';
  }

  /* =============================
     Helpers de fechas
     ============================= */
  function isDateHeader(h){
    return /fecha/i.test(h);
  }

  function toInputDateValue(raw, type){
    const s = String(raw || '').trim();
    if (!s) return '';

    // DD/MM/YYYY o DD-MM-YYYY (+ opcional HH:MM)
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T](\d{1,2}):(\d{2}))?$/);
    if (m){
      const dd = String(m[1]).padStart(2,'0');
      const mm = String(m[2]).padStart(2,'0');
      const yyyy = m[3].length === 2 ? ('20' + m[3]) : m[3];
      const date = `${yyyy}-${mm}-${dd}`;

      if (type === 'datetime-local' && m[4] && m[5]){
        const hh = String(m[4]).padStart(2,'0');
        const mi = String(m[5]).padStart(2,'0');
        return `${date}T${hh}:${mi}`;
      }
      return date;
    }

    // Ya viene ISO
    if (/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2})?$/.test(s)) return s;

    return s;
  }

  function todayISODate() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  function currentDateTimeLocal() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const hh = String(d.getHours()).padStart(2, '0');
    const mi = String(d.getMinutes()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}T${hh}:${mi}`;
  }

  /* =============================
     Select helpers
     ============================= */
  function selectHTML(id, header, options, current, disabled, dataAttrs){
    const attrs = [
      `id="${id}"`,
      `name="${escAttr(header)}"`,
      `class="in control"`,
      disabled ? 'disabled' : ''
    ];
    if (dataAttrs && dataAttrs.role) attrs.push(`data-role="${dataAttrs.role}"`);
    if (dataAttrs && dataAttrs.idx != null) attrs.push(`data-idx="${String(dataAttrs.idx)}"`);

    return [
      `<select ${attrs.join(' ')}>`,
      selectOptionsHTML(options, current),
      `</select>`
    ].join('');
  }

  function selectOptionsHTML(options, current){
    options = Array.isArray(options) ? options : [];
    const hasCurrent = current && options.includes(current);

    const optHtml = options.map(opt =>
      `<option value="${escAttr(opt)}" ${opt===current?'selected':''}>${esc(opt)}</option>`
    ).join('');

    return [
      `<option value="" ${(!current || !hasCurrent) ? 'selected' : ''}>— Selecciona —</option>`,
      optHtml
    ].join('');
  }

  /* =============================
     Sorting helpers
     ============================= */
  function compareByKey(a, b, key) {
    const A = normalizeForSort(a?.[key]);
    const B = normalizeForSort(b?.[key]);

    const aNull = (A === null || A === undefined || A === '');
    const bNull = (B === null || B === undefined || B === '');
    if (aNull && bNull) return 0;
    if (aNull) return 1;
    if (bNull) return -1;

    if (A instanceof Date && B instanceof Date) return A - B;
    if (typeof A === 'number' && typeof B === 'number') return A - B;

    return String(A).localeCompare(String(B), 'es', { sensitivity: 'base', ignorePunctuation: true });
  }

  function normalizeForSort(v) {
    if (v == null) return '';
    const s = String(v).trim();

    const d = parseYMD(s);
    if (d instanceof Date && !isNaN(d)) return d;

    const numericish = s.replace(/[^\d.\-]/g, '');
    if (numericish && /^-?\d+(\.\d+)?$/.test(numericish)) {
      const n = Number(numericish);
      if (!isNaN(n)) return n;
    }

    return s.toLowerCase();
  }

  /* =============================
     7) Alertas / Urgencias (si existe #alerts en el HTML)
     ============================= */
  async function computeAndRenderAlerts() {
    if (!alertsEl) return;

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
      if (u.type === 'urgent' || u.type === 'today') urgent.push({ row:r, when:u.whenSort || null });
      else if (u.type === 'soon') soon.push({ row:r, when:u.whenSort || null });
    });

    const byWhen = (a,b) => {
      if (a.when && b.when && a.when.getTime() !== b.when.getTime()) return a.when < b.when ? -1 : 1;
      const ap = (String(a.row['Prioridad']||'').toLowerCase()==='alta') ? 0 : 1;
      const bp = (String(b.row['Prioridad']||'').toLowerCase()==='alta') ? 0 : 1;
      return ap - bp;
    };

    urgent.sort(byWhen);
    soon.sort(byWhen);

    renderAlertsUI(
      urgent.slice(0,8).map(x=>x.row),
      soon.slice(0,4).map(x=>x.row),
      urgent.length,
      soon.length
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
      const urgency = classifyUrgency(r);
      const name  = (String(r['Nombre']||'').trim()) || '(Sin nombre)';
      const prio  = (String(r['Prioridad']||'').trim()) || '';
      const canal = (String(r['Canal de comunicación'] || r['Canal'] || '').trim());
      const when  = minDate(parseYMD(r['Fecha para contactar']), parseYMD(r['Siguiente Contacto (calc)']));
      const whenTxt = urgency.reason || (when ? formatYMD(when) : 'Sin fecha');
      const id = String(r['ID']||'').trim();

      const badge = prio ? prio : 'Pendiente';

      return `
        <div class="alert-item">
          <div class="alert-person">
            <strong>${esc(name)}</strong>
            <span class="meta">${canal ? esc(canal) : 'Sin canal'} · ${esc(whenTxt)}</span>
          </div>
          <span class="mini-priority">${esc(badge)}</span>
          <button class="btn alert-open" data-openid="${escAttr(id)}">Abrir</button>
        </div>
      `;
    };

    alertsEl.innerHTML = `
      <div class="alert-card">
        <div class="alert-title">
          <div>
            <span class="alert-eyebrow">Seguimiento</span>
            <h2>Contactos que necesitan atención</h2>
          </div>
          <div class="pills">
            <span class="pill urgent">${uCount.toLocaleString()} urgentes</span>
            <span class="pill soon">${sCount.toLocaleString()} próximos</span>
          </div>
        </div>
        <div class="alert-list">
          ${urgentTop.map(makeItem).join('')}
          ${soonTop.map(makeItem).join('')}
        </div>
      </div>
    `;
    alertsEl.classList.remove('hidden');

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

  function classifyUrgency(row){
    const prio = (String(row['Prioridad'] ?? '')).trim().toLowerCase();
    const f1   = parseYMD(String(row['Fecha para contactar'] ?? ''));
    const f2   = parseYMD(String(row['Siguiente Contacto (calc)'] ?? ''));
    const when = minDate(f1, f2);
    const lastContact = getLastContactDate(row);
    const enrolled = isEnrolled(row);

    const today = toYMD(new Date());
    const in3   = addDays(today, 3);

    const dueOrPast = (d) => d && d <= today;
    const within3   = (d) => d && d > today && d <= in3;

    if (enrolled) return { type:'ok', when:'', whenSort: null, reason:'Matriculado' };
    if (prio === 'alta' && dueOrPast(when)) return { type:'urgent', when: formatIf(when), whenSort: when, reason:'Contacto vencido' };
    if (prio === 'alta' && !when)           return { type:'urgent', when: '', whenSort: today, reason:'Prioridad alta sin fecha' };
    if (dueOrPast(when))                    return { type: when && when.getTime() === today.getTime() ? 'today' : 'urgent', when: formatIf(when), whenSort: when || today, reason:'Fecha de contacto vencida' };
    if (lastContact) {
      const days = daysBetween(lastContact, today);
      if (days >= 30) return { type:'urgent', when: `${days} dias`, whenSort: addDays(today, -days), reason:'Sin contacto hace 1 mes o mas' };
      if (days >= 15) return { type:'urgent', when: `${days} dias`, whenSort: addDays(today, -days), reason:'Sin contacto hace 15 dias' };
      if (days >= 7)  return { type:'soon', when: `${days} dias`, whenSort: addDays(today, -days), reason:'Sin contacto hace 1 semana' };
    } else {
      return { type:'urgent', when:'', whenSort: today, reason:'Nunca contactado' };
    }
    if (within3(when))                      return { type:'soon', when: formatIf(when), whenSort: when, reason:'Contacto cercano' };
    return { type:'ok', when:'', whenSort: null };
  }

  function getLastContactDate(row) {
    const candidates = [
      'Último contacto', 'Último Contacto', 'Ultimo contacto', 'Ultimo Contacto',
      'Fecha y hora de contacto', 'Fecha de contacto', 'Fecha Contacto',
      'Fecha para contactar'
    ];
    for (const key of candidates) {
      const parsed = parseYMD(String(row[key] ?? ''));
      if (parsed) return parsed;
    }
    return null;
  }

  function isEnrolled(row) {
    const blob = Object.keys(row || {})
      .filter(k => /estado|matric|listado|status/i.test(k))
      .map(k => String(row[k] ?? '').toLowerCase())
      .join(' | ');
    return /matriculad|inscrit|activo/.test(blob) && !/no\s+matriculad|no\s+inscrit|retirad/.test(blob);
  }

  async function getRowById(id) {
    const idKey = keyColumnInUse || 'ID';

    // primero en lo visible
    let row = displayedRows.find(r => String(r[idKey] ?? '').trim() === id);
    if (row) return row;

    // luego en cache completo
    let cache = allRowsCache.get(currentSheet);
    if (!cache) {
      cache = await loadAllRowsForSheet(currentSheet);
      allRowsCache.set(currentSheet, cache);
    }
    row = (cache.rows || []).find(r => String(r[idKey] ?? '').trim() === id);
    return row || null;
  }

  /* =============================
     Helpers fecha (reutilizados)
     ============================= */
  function toYMD(d){
    return d instanceof Date ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null;
  }
  function addDays(d, n){
    const x = new Date(d);
    x.setDate(x.getDate()+n);
    return toYMD(x);
  }
  function daysBetween(a,b){
    if (!a || !b) return 0;
    return Math.floor((toYMD(b) - toYMD(a)) / 86400000);
  }
  function minDate(a,b){
    if (a && b) return a < b ? a : b;
    return a || b || null;
  }
  function parseYMD(s){
    if (!s) return null;
    const str = String(s).trim();

    const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return toYMD(new Date(Number(iso[1]), Number(iso[2])-1, Number(iso[3])));

    const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T]\d{1,2}:\d{2})?/);
    if (dmy){
      const dd = Number(dmy[1]);
      const mm = Number(dmy[2]) - 1;
      const yyyy = dmy[3].length === 2 ? Number('20' + dmy[3]) : Number(dmy[3]);
      return toYMD(new Date(yyyy, mm, dd));
    }
    return null;
  }
  function pad2(n){ return String(n).padStart(2,'0'); }
  function formatYMD(d){ return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; }
  function formatIf(d){ return d ? formatYMD(d) : ''; }

  // ✅ cerrar IIFE
})();
