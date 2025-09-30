/* =============================
   app.js — Callcenter Musicala
   Paginación + búsqueda + modal editable
   Actualiza con POST (mode=update) y Crea con POST (mode=add)
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
  const btnPrev    = document.getElementById('btnPrev');
  const btnNext    = document.getElementById('btnNext');
  const pageInfo   = document.getElementById('pageInfo');
  const pageSizeEl = document.getElementById('pageSize');

  // Modal
  const modal         = document.getElementById('clientModal');
  const modalBody     = document.getElementById('clientBody');
  const modalTitle    = document.getElementById('clientTitle');
  const modalSubtitle = document.getElementById('clientSubtitle');

  // ------- Config -------
  if (!window.API_BASE) throw new Error('No se encontró window.API_BASE. Define la URL /exec en index.html');
  const API_BASE = window.API_BASE;

  // Detectaremos automáticamente la columna clave según headers disponibles
  const KEY_CANDIDATES = ['ID', 'Listado1', 'Celular/Teléfono'];

  // ====== Catálogos (listas fijas) ======
  const OPTIONS = {
    Grupo: ["Musiadultos","Musibabies","Musicalitos","Musigrandes","Musikids","Musiteens"],

    // lista para Arte I, II y III
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

  // ====== Dependencias de Instrumento/Estilo/Técnica (según Arte) ======
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

  // Campos Arte / Instrumento por índice
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
  let keyColumnInUse = ''; // columna clave efectiva de la hoja cargada
  let creatingNew    = false; // bandera "nuevo"

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
    await loadPage();
  });

  btnPrev?.addEventListener('click', async () => {
    if (offset <= 0) return;
    offset = Math.max(0, offset - limit);
    await loadPage();
  });

  btnNext?.addEventListener('click', async () => {
    if (offset + limit >= total) return;
    offset = offset + limit;
    await loadPage();
  });

  pageSizeEl?.addEventListener('change', async () => {
    limit = parseInt(pageSizeEl.value, 10) || 200;
    offset = 0;
    await loadPage();
  });

  searchBox.addEventListener('input', filterTable);

  // Nuevo cliente
  btnNew?.addEventListener('click', () => openNewModal());

  // Delegación: click en el "Nombre" abre modal de edición
  table.addEventListener('click', (e) => {
    const btn = e.target.closest('.name-link');
    if (!btn) return;
    const idx = Number(btn.dataset.rowindex);
    const row = currentRows[idx];
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
     3) Lógica de datos
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

    // Elegir columna clave efectiva según headers
    keyColumnInUse = pickKeyColumn(currentHeaders);
    if (!keyColumnInUse) {
      status('⚠️ No encontré una columna clave (ID/Listado1/Celular/Teléfono). Podrás ver datos, pero no guardar.');
    } else {
      status(`Listo. Clave: “${keyColumnInUse}”. Mostrando ${currentRows.length.toLocaleString()} de ${total.toLocaleString()} registros.`);
    }

    renderTable(currentHeaders, currentRows);

    const page  = Math.floor(offset / limit) + 1;
    const pages = Math.max(1, Math.ceil(total / limit));
    if (pageInfo) {
      pageInfo.textContent = `Página ${page.toLocaleString()} de ${pages.toLocaleString()} — ${total.toLocaleString()} registros`;
    }
    if (btnPrev) btnPrev.disabled = offset <= 0;
    if (btnNext) btnNext.disabled = offset + limit >= total;

    filterTable();
  }

  function renderTable(headers, rows) {
    const nameIdx = headers.findIndex((h) => h.trim().toLowerCase() === 'nombre');

    const thead = `<thead><tr>${
      headers.map((h) => `<th>${esc(h)}</th>`).join('')
    }</tr></thead>`;

    const tbody = `<tbody>${
      rows.map((r, i) => {
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

    const pxPerCol = 140;
    table.style.minWidth = (headers.length * pxPerCol) + 'px';
  }

  /* =============================
     4) Modales
     ============================= */

  // Abre modal para fila existente (editar)
  function openModal(row) {
    creatingNew = false;
    buildModal(row);
  }

  // Abre modal vacío para crear
  function openNewModal() {
    creatingNew = true;

    // armar objeto vacío con todos los headers conocidos
    const empty = {};
    currentHeaders.forEach(h => empty[h] = '');

    // opcional: si existe columna ID, generamos uno temporal para no perder la clave
    if (currentHeaders.includes('ID')) {
      empty['ID'] = genId(); // puedes quitarlo si prefieres escribirlo manualmente
    }
    // si quieres default del mes actual en Listado/Listado1:
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
    modalTitle.textContent = nombre;

    const canal = String(row['Canal de comunicación'] ?? '').trim();
    const grupo = String(row['Grupo'] ?? '').trim();
    const arte1 = String(row['Arte I'] ?? '').trim();
    modalSubtitle.textContent = [grupo, arte1, canal].filter(Boolean).join(' · ');

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

      // Selects fijos
      if (OPTIONS[headerName]) {
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
          // Validaciones mínimas: clave si existe / nombre opcional
          if (keyColumnInUse) {
            const k = String(formValues[keyColumnInUse] || '').trim();
            if (!k) throw new Error(`Falta valor en la columna clave: ${keyColumnInUse}`);
          }
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

        // Edición existente
        const changes = diffObject(originalRow, formValues);
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
          upsert: 'true',
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
     5) Utilidades
     ============================= */
  function filterTable() {
    const q = searchBox.value.trim().toLowerCase();
    table.querySelectorAll('tbody tr').forEach((tr) => {
      tr.style.display = !q || tr.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
  }

  function pickKeyColumn(headers) {
    for (const k of KEY_CANDIDATES) if (headers.includes(k)) return k;
    return ''; // no hay clave conocida
  }

  function genId(){
    // ID corto y único (fecha + random base36)
    return 'ID-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2,6).toUpperCase();
  }

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
})();
