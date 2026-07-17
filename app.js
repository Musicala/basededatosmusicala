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
  const btnAdvisorStats = document.getElementById('btnAdvisorStats');
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
  const todayViewEl   = document.getElementById('todayView');
  const advisorStatsPanel = document.getElementById('advisorStatsPanel');
  const advisorStatsContent = document.getElementById('advisorStatsContent');
  const advisorStatsSummary = document.getElementById('advisorStatsSummary');
  const statsFrom = document.getElementById('statsFrom');
  const statsTo = document.getElementById('statsTo');
  const btnRefreshStats = document.getElementById('btnRefreshStats');
  const btnCloseStats = document.getElementById('btnCloseStats');

  // Conciliación con Lista de estudiantes (proyecto Firebase separado)
  const btnReconcile = document.getElementById('btnReconcile');
  const reconcilePanel = document.getElementById('reconcilePanel');
  const reconcileSummary = document.getElementById('reconcileSummary');
  const reconcileContent = document.getElementById('reconcileContent');
  const btnRunReconcile = document.getElementById('btnRunReconcile');
  const btnSaveReconcile = document.getElementById('btnSaveReconcile');
  const btnCloseReconcile = document.getElementById('btnCloseReconcile');

  // ------- Config -------
  // Sheets (Apps Script) es SOLO respaldo opcional: la app funciona 100% con Firebase.
  const API_BASE = window.API_BASE || '';
  const APP_VERSION = '1.5.0';
  const appVersionEl = document.getElementById('appVersion');
  if (appVersionEl) appVersionEl.textContent = `v${APP_VERSION}`;
  const ALLOWED_EMAILS = [
    'alekcaballeromusic@gmail.com',
    'catalina.medina.leal@gmail.com',
    'musicalaasesor@gmail.com',
    'imusicala@gmail.com'
  ];
  const ADMIN_EMAILS = [
    'alekcaballeromusic@gmail.com',
    'catalina.medina.leal@gmail.com'
  ];
  const firebaseDb = initFirebaseCache();
  const authReady = waitForFirebaseAuth();

  // 🔒 Clave SOLO "ID"
  const KEY_CANDIDATES = ['ID'];

  const FIELD_ALIASES = {
    id: ['ID'],
    name: ['Nombre', 'Nombre completo', 'Nombre acudiente', 'Estudiante'],
    phone: ['Celular/TelÃ©fono', 'Celular/Telefono', 'Telefono', 'TelÃ©fono', 'Celular', 'WhatsApp'],
    email: ['Correo ElectrÃ³nico', 'Correo Electronico', 'Email', 'Correo'],
    channel: ['Canal de comunicaciÃ³n', 'Canal de comunicacion', 'Canal'],
    advisor: ['Asesor', 'Responsable', 'Asesor asignado'],
    art: ['ARTE', 'Arte I', 'Arte principal'],
    instrument: ['Instrumento/Estilo/TÃ©cnica I', 'Instrumento/Estilo/Tecnica I', 'Instrumento', 'Estilo'],
    modality: ['Modalidad'],
    plan: ['Curso/Plan', 'Plan', 'Curso'],
    priority: ['Prioridad'],
    nextContact: ['Fecha para contactar', 'Proximo contacto', 'PrÃ³ximo contacto', 'Fecha proximo contacto', 'Fecha prÃ³ximo contacto'],
    lastContact: ['Fecha y hora de contacto', 'Fecha ultima gestion', 'Fecha Ãºltima gestiÃ³n', 'Ultima gestion', 'Ãšltima gestiÃ³n'],
    lastManaged: ['Fecha ultima gestion', 'Fecha Ãºltima gestiÃ³n', 'Ultima gestion', 'Ãšltima gestiÃ³n'],
    result: ['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado', 'Resultado de gestion', 'Resultado de gestiÃ³n'],
    status: ['Estado de seguimiento', 'Estado', 'Status'],
    attempts: ['Intentos', 'Numero de intentos', 'NÃºmero de intentos'],
    exclusionReason: ['Motivo de exclusion', 'Motivo de exclusiÃ³n'],
    nextAudit: ['Fecha proxima auditoria', 'Fecha prÃ³xima auditorÃ­a', 'Proxima auditoria', 'PrÃ³xima auditorÃ­a'],
    lastRevision: ['Fecha ultima revision', 'Fecha Ãºltima revisiÃ³n']
  };

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
  const searchBlobCache = new WeakMap(); // fila -> texto normalizado para buscar
  const MAX_RENDER_ROWS = 1000; // tope de filas pintadas de una vez (evita congelar la pestaña)
  const allRowsLoadPromises = new Map(); // sheet -> Promise<{ headers, rows, total, ts }>
  let searchActive   = false;

  // OJO: displayedRows SIEMPRE debe ser exactamente lo que está renderizado
  let displayedRows  = []; // filas actualmente renderizadas (página o filtradas y/o ordenadas)
  let alertFilter = 'all';
  let todayViewExpanded = false;
  let followUpRefreshTimer = null;
  let followUpRefreshToken = 0;
  let dailySyncScheduled = false;
  let fauxProgressTimer = null;
  let hideProgressTimer = null;
  const TOP_RECOMMENDED_LIMIT = 10;
  const NO_RESPONSE_COOLDOWN_DAYS = 2;
  const CONTACTED_COOLDOWN_DAYS = 7;
  const SKIPPED_COOLDOWN_DAYS = 1;
  const EXCLUDED_AUDIT_DAYS = 180;
  const CONFIRMED_EXCLUDED_AUDIT_DAYS = 365;
  const TOP_RECOMMENDED_QUOTAS = {
    new_recent: 2,
    due_recent: 2,
    high_no_date: 1,
    reactivation_never_contacted: 2,
    reactivation_old_client_or_warm: 1,
    excluded_audit: 1,
    general_rotation: 1
  };
  const CRM_STAGES = {
    new: { label: 'Nuevo', className: 'crm-new' },
    contacted: { label: 'Contactado', className: 'crm-contacted' },
    interested: { label: 'Interesado', className: 'crm-interested' },
    trial_scheduled: { label: 'Clase agendada', className: 'crm-trial' },
    trial_done: { label: 'Clase realizada', className: 'crm-trial' },
    payment_pending: { label: 'Pendiente pago', className: 'crm-payment' },
    enrolled: { label: 'Matriculado', className: 'crm-enrolled' },
    not_interested: { label: 'No interesado', className: 'crm-closed' },
    reactivation: { label: 'Reactivar', className: 'crm-reactivation' },
    invalid: { label: 'Datos invalidos', className: 'crm-invalid' },
    archived: { label: 'Archivado', className: 'crm-archived' }
  };
  const CRM_TEMPERATURES = {
    hot: { label: 'Caliente', className: 'temp-hot' },
    warm: { label: 'Tibio', className: 'temp-warm' },
    cold: { label: 'Frio', className: 'temp-cold' },
    frozen: { label: 'Congelado', className: 'temp-frozen' }
  };
  const WHATSAPP_TEMPLATES = [
    { id:'first_contact', label:'Primer contacto', text:'Hola, {{nombre}}. Te saluda {{asesor}} de Musicala. Vimos que estas interesado/a en {{arte}} y queremos ayudarte a encontrar el mejor horario para ti. Te puedo compartir opciones en {{modalidad}}?' },
    { id:'kind_followup', label:'Seguimiento amable', text:'Hola, {{nombre}}. Paso a saludarte desde Musicala para saber si pudiste revisar la informacion de {{arte}}. Quieres que miremos horarios o modalidad?' },
    { id:'no_response', label:'No respuesta', text:'Hola, {{nombre}}. Te escribo nuevamente de Musicala. Si aun te interesa {{arte}}, puedo ayudarte a encontrar una opcion que se acomode a tu tiempo.' },
    { id:'trial_confirm', label:'Confirmacion clase de prueba', text:'Hola, {{nombre}}. Confirmamos tu clase de prueba de {{arte}} para el {{fecha}} a las {{hora}}. Modalidad: {{modalidad}}. Te esperamos en Musicala.' },
    { id:'trial_reminder', label:'Recordatorio clase de prueba', text:'Hola, {{nombre}}. Te recordamos tu clase de prueba de {{arte}} programada para {{fecha}} a las {{hora}}. Cualquier duda me escribes.' },
    { id:'after_trial', label:'Despues de clase de prueba', text:'Hola, {{nombre}}. Gracias por vivir la clase de prueba de {{arte}} con Musicala. Quieres que avancemos con horarios y proceso de matricula?' },
    { id:'payment_pending', label:'Pendiente de pago', text:'Hola, {{nombre}}. Quedamos atentos para ayudarte a finalizar el proceso de pago y separar tu cupo de {{arte}} en Musicala.' },
    { id:'reactivation', label:'Reactivacion', text:'Hola, {{nombre}}. Te saluda {{asesor}} de Musicala. Queremos saber si te gustaria retomar tu interes por {{arte}}. Tenemos opciones nuevas que podrian servirte.' },
    { id:'vacation', label:'Vacacionales', text:'Hola, {{nombre}}. En Musicala tenemos programas vacacionales para disfrutar y aprender arte. Te gustaria que te comparta opciones?' },
    { id:'adults', label:'Adultos', text:'Hola, {{nombre}}. En Musicala tambien tenemos clases para adultos en {{arte}}. Podemos revisar una opcion segun tu nivel y disponibilidad.' },
    { id:'music', label:'Musica', text:'Hola, {{nombre}}. Te comparto informacion sobre clases de musica en Musicala. Podemos revisar instrumento, modalidad y horarios.' },
    { id:'dance', label:'Danzas', text:'Hola, {{nombre}}. Tenemos opciones de danza en Musicala para diferentes edades y niveles. Te gustaria que revisemos horarios?' },
    { id:'theatre', label:'Teatro', text:'Hola, {{nombre}}. En Musicala tenemos clases de teatro para explorar expresion, creatividad y confianza. Te comparto opciones?' },
    { id:'visual_arts', label:'Artes plasticas', text:'Hola, {{nombre}}. Tenemos clases de artes plasticas en Musicala. Podemos revisar tecnica, edad y modalidad para recomendarte la mejor opcion.' },
    { id:'closing', label:'Cierre por no respuesta', text:'Hola, {{nombre}}. Como no hemos logrado comunicarnos, dejaremos tu solicitud en pausa por ahora. Si quieres retomar {{arte}} en Musicala, puedes escribirme y con gusto te ayudo.' }
  ];
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
  setupAdvisorAdminUI();

  const sheetNames = await loadSheetNamesFromFirebase();
  if (!sheetNames.length) {
    sheetSel.innerHTML = '<option>Sin datos en Firebase</option>';
    currentHeaders = [];
    currentRows = [];
    total = 0;
    renderTable([], []);
    status('No hay hojas cargadas en Firebase. Usa "Actualizar datos" despues de elegir una hoja inicial en la configuracion o haz una primera sincronizacion desde Sheets.');
    setFirebaseStatus('No hay cache de hojas en Firebase todavia.');
    return;
  }

  sheetSel.innerHTML = sheetNames
    .map((n) => `<option value="${escAttr(n)}">${esc(n)}</option>`)
    .join('');

  const initial = sheetNames.find((n) => /base de datos general/i.test(n)) || sheetNames[0];
  sheetSel.value = initial;
  currentSheet   = initial;

  await loadPage(true);

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
    await loadPage(false, false);
  });

  btnNext?.addEventListener('click', async () => {
    if (searchActive) return;
    if (offset + limit >= total) return;
    offset = offset + limit;
    await loadPage(false, false);
  });

  pageSizeEl?.addEventListener('change', async () => {
    if (searchActive) return;
    limit = parseInt(pageSizeEl.value, 10) || 200;
    offset = 0;
    await loadPage(false, false);
  });

  // Búsqueda GLOBAL con debounce
  let debounce;
  searchBox.addEventListener('input', () => {
    clearTimeout(debounce);
    debounce = setTimeout(applySearch, 250);
  });

  // Nuevo cliente
  btnNew?.addEventListener('click', () => openNewModal());
  btnAdvisorStats?.addEventListener('click', () => openAdvisorStats());
  btnRefreshStats?.addEventListener('click', () => renderAdvisorStats());
  btnCloseStats?.addEventListener('click', () => advisorStatsPanel?.classList.add('hidden'));
  btnSyncFirebase?.addEventListener('click', async () => {
    try {
      btnSyncFirebase.disabled = true;
      await syncCurrentSheetToFirebase(true);
      await logAdvisorActivity('sync_firebase', {
        sheetName: currentSheet,
        label: 'Actualizo datos',
        count: total || currentRows.length || 0
      });
    } finally {
      btnSyncFirebase.disabled = false;
    }
  });
  btnSyncSheets?.addEventListener('click', async () => {
    try {
      btnSyncSheets.disabled = true;
      const synced = await syncPendingFirebaseToSheets(true);
      await logAdvisorActivity('sync_sheets', {
        sheetName: currentSheet,
        label: 'Guardo en principal',
        count: synced || 0
      });
    } finally {
      btnSyncSheets.disabled = false;
    }
  });
  btnReconcile?.addEventListener('click', () => {
    reconcilePanel?.classList.remove('hidden');
    reconcilePanel?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  btnCloseReconcile?.addEventListener('click', () => reconcilePanel?.classList.add('hidden'));
  btnRunReconcile?.addEventListener('click', async () => {
    try {
      btnRunReconcile.disabled = true;
      await runReconciliation();
    } catch (err) {
      console.error(err);
      if (reconcileSummary) reconcileSummary.textContent = `No se pudo conciliar: ${err?.message || err}`;
    } finally {
      btnRunReconcile.disabled = false;
    }
  });
  btnSaveReconcile?.addEventListener('click', async () => {
    try {
      btnSaveReconcile.disabled = true;
      await saveReconciliationToFirebase();
    } catch (err) {
      console.error(err);
      if (reconcileSummary) reconcileSummary.textContent = `No se pudo guardar la conciliación: ${err?.message || err}`;
    } finally {
      btnSaveReconcile.disabled = false;
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
  async function loadPage(withProgress = false, refreshToday = true) {
    status(`Cargando “${currentSheet}”…`);
    if (withProgress) startFauxProgress('Cargando la base de datos', `Conectando con la base y trayendo "${currentSheet}"...`, { start: 12, ceiling: 85 });
    // Vigilante: si la base no responde, avisamos en vez de quedarnos colgados en silencio.
    const watchdog = window.setTimeout(() => {
      status('La base esta tardando mas de lo normal. Revisa tu conexion a internet; si persiste, recarga la pagina (Ctrl+Shift+R).');
      if (withProgress) setProgress(86, 'Esperando respuesta de la base...');
    }, 15000);
    // La hoja completa se descarga UNA sola vez y queda en cache en memoria
    // (y en el cache local del navegador). Paginar, buscar y "Hoy" reutilizan
    // esa misma copia: Anterior/Siguiente ya no consultan la red.
    let cached;
    try {
      cached = await loadAllRowsForSheet(currentSheet);
      if (!cached?.rows?.length) {
        // Respaldo: lectura paginada directa (comportamiento anterior).
        cached = await loadSheetPageFromFirebase(currentSheet, offset, limit);
      }
    } catch (err) {
      window.clearTimeout(watchdog);
      if (withProgress) { stopFauxProgressTimer(); hideProgress(); }
      status(`No se pudo cargar la base: ${err?.message || err}. Recarga la pagina (Ctrl+Shift+R).`);
      throw err;
    }
    window.clearTimeout(watchdog);
    if (withProgress) finishFauxProgress('Base lista.');
    if (cached?.rows?.length) {
      currentHeaders = cached.headers || [];
      total = Number(cached.total || cached.rows.length);
      if (offset >= total) offset = Math.max(0, Math.floor(Math.max(0, total - 1) / limit) * limit);
      currentRows = cached.source === 'firebase-page'
        ? (cached.rows || [])
        : (cached.rows || []).slice(offset, offset + limit);
      keyColumnInUse = pickKeyColumn(currentHeaders);
      renderTable(currentHeaders, currentRows);

      const page  = Math.floor(offset / limit) + 1;
      const pages = Math.max(1, Math.ceil(total / limit));
      if (pageInfo) pageInfo.textContent = `Página ${page.toLocaleString()} de ${pages.toLocaleString()} — ${total.toLocaleString()} registros`;
      if (btnPrev) btnPrev.disabled = offset <= 0;
      if (btnNext) btnNext.disabled = offset + limit >= total;
      status(`Listo. Mostrando ${currentRows.length.toLocaleString()} de ${total.toLocaleString()} registros.`);
      // "Hoy" depende de toda la base, no de la pagina visible: solo lo refrescamos
      // en carga inicial / cambio de hoja / cambios de datos, no al paginar.
      if (refreshToday) {
        renderFollowUpPlaceholder();
        scheduleFollowUpRefresh(250);
      }
      return;
    }

    currentHeaders = [];
    currentRows = [];
    total = 0;
    keyColumnInUse = '';
    renderTable(currentHeaders, currentRows);
    if (pageInfo) pageInfo.textContent = 'Sin registros en Firebase';
    if (btnPrev) btnPrev.disabled = true;
    if (btnNext) btnNext.disabled = true;
    status(`No hay datos guardados en Firebase para "${currentSheet}". Usa "Actualizar datos" para traerlos desde Sheets.`);
    setFirebaseStatus('Modo Firebase-first: Sheets solo se consulta con Actualizar datos.');
    renderFollowUpPlaceholder();
    scheduleFollowUpRefresh(250);
  }

  function renderTable(headers, rows) {
    const nameIdx = headers.findIndex((h) => h.trim().toLowerCase() === 'nombre');
    const visibleHeaders = ['__crmStage', '__crmTemperature', '__enrollment', ...headers];

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

    // Tope de render: pintar miles de filas de golpe congela la pestaña.
    const truncated = toRender.length > MAX_RENDER_ROWS;
    if (truncated) toRender = toRender.slice(0, MAX_RENDER_ROWS);

    // ✅ LO MÁS IMPORTANTE: lo que se ve = displayedRows
    displayedRows = toRender;

    const thead = `<thead><tr>${
      visibleHeaders.map((h) => {
        const isActive = sortState.key === h;
        const dirAttr  = isActive ? ` data-dir="${sortState.dir}"` : '';
        const label = h === '__crmStage' ? 'Etapa CRM' : h === '__crmTemperature' ? 'Temperatura' : h === '__enrollment' ? 'Inscrito' : h;
        return `<th class="th-sortable" data-key="${escAttr(h)}"${dirAttr}>
                  <span>${esc(label)}</span>
                  <span class="th-sort-ind"></span>
                </th>`;
      }).join('')
    }</tr></thead>`;

    const tbody = `<tbody>${
      toRender.map((r, i) => {
        const crm = normalizeCrm(r);
        const tds = visibleHeaders.map((h) => {
          if (h === '__crmStage') return `<td>${renderCrmStageBadge(crm.stage)}</td>`;
          if (h === '__crmTemperature') return `<td>${renderTemperatureBadge(crm.temperature, crm.leadScore)}</td>`;
          if (h === '__enrollment') return `<td>${renderEnrollmentBadge(r.__enrollment)}</td>`;
          const colIdx = headers.indexOf(h);
          const val = r[h];
          if (colIdx === nameIdx) {
            const label = (val == null || String(val).trim() === '') ? '(Sin nombre)' : String(val);
            return `<td><button class="linkish name-link" data-rowindex="${i}">${esc(label)}</button></td>`;
          }
          return `<td>${pretty(val)}</td>`;
        }).join('');
        return `<tr>${tds}</tr>`;
      }).join('')
    }${truncated ? `<tr><td colspan="${visibleHeaders.length}" class="locked-cell">Mostrando las primeras ${MAX_RENDER_ROWS.toLocaleString()} filas. Usa la búsqueda para afinar el resultado.</td></tr>` : ''}</tbody>`;

    table.innerHTML = thead + tbody;

    // ancho mínimo para scroll
    const pxPerCol = 140;
    table.style.minWidth = (visibleHeaders.length * pxPerCol) + 'px';

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
    const cachedInMemory = allRowsCache.get(sheetName);
    if (cachedInMemory?.rows?.length) return cachedInMemory;

    const pending = allRowsLoadPromises.get(sheetName);
    if (pending) return pending;

    const loadPromise = (async () => {
      const cached = await loadSheetFromFirebase(sheetName);
      if (cached) return cached;

      return { headers: currentHeaders, rows: [], total: 0, ts: Date.now(), source: 'firebase-empty' };
    })();

    allRowsLoadPromises.set(sheetName, loadPromise);
    try {
      const loaded = await loadPromise;
      allRowsCache.set(sheetName, loaded);
      return loaded;
    } finally {
      allRowsLoadPromises.delete(sheetName);
    }
  }

  async function fetchAllRowsFromSheets(sheetName, mirrorToFirebase = false) {
    if (!API_BASE) throw new Error('El respaldo de Google Sheets no está configurado (window.API_BASE).');
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

      // El texto normalizado de cada fila se calcula UNA sola vez y se reutiliza
      // en cada tecleo: así la búsqueda no congela la página con bases grandes.
      const filtered = rows.filter(r => {
        let blob = searchBlobCache.get(r);
        if (blob === undefined) {
          blob = norm(headers.map(h => r[h]).join(' | '));
          searchBlobCache.set(r, blob);
        }
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
    modalBody.innerHTML = renderCrmSummaryHTML(row) + renderWhatsAppTemplatesHTML(row) + renderHistoryHTML(isNew) + headers.map((h) => renderFieldHTML(h, row[h] ?? '', row, isNew)).join('');

    // ✅ Ahora sí: como ya existe el DOM del modal, conectamos WhatsApp + dependencias
    wireWhatsAppButton();
    wireWhatsAppTemplates(row);
    if (!isNew) loadModalHistory(row);
    if (isNew) wireArteDependencies(); // en nuevo, ya editable

    const footer = modal.querySelector('.modal-footer');
    footer.innerHTML = `
      <span id="saveStatus" class="muted"></span>
      <div style="margin-left:auto; display:flex; gap:.5rem;">
        <button id="btnContactNow" class="btn" type="button">Registrar contacto ahora</button>
        ${isNew ? '' : '<button id="btnEdit" class="btn">Editar</button>'}
        <button id="btnSave" class="btn">${isNew ? 'Crear' : 'Guardar'}</button>
        <button class="btn" data-close="1">Cerrar</button>
      </div>`;

    const btnEdit    = document.getElementById('btnEdit');
    const btnSave    = document.getElementById('btnSave');
    const btnContactNow = document.getElementById('btnContactNow');
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

    btnContactNow?.addEventListener('click', () => {
      const field = modalBody.querySelector('[data-contact-real="1"]');
      if (!field) {
        saveStatus.textContent = 'No hay campo de fecha de contacto real en esta hoja.';
        return;
      }
      field.removeAttribute('readonly');
      field.removeAttribute('disabled');
      field.value = currentDateTimeLocal();
      btnSave.disabled = false;
      saveStatus.textContent = 'Contacto marcado para guardar';
    });

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
          if (!String(formValues.ID || '').trim()) formValues.ID = `crm-${Date.now()}`;
          await saveRowToFirebase(currentSheet, formValues, true);
          saveStatus.textContent = 'Creado en Firebase. Pendiente de enviar a principal';
          await logAdvisorActivity('create_record', {
            sheetName: currentSheet,
            rowId: String(formValues.ID || ''),
            label: 'Creo registro',
            contactSnapshot: buildContactSnapshot(formValues)
          });
          // Actualizamos la copia local en vez de re-descargar toda la base.
          upsertRowInLocalState(formValues);
          await loadPage(false, false);
          scheduleFollowUpRefresh(300);
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

        const mergedRow = { ...originalRow, ...formValues };
        await saveRowToFirebase(currentSheet, mergedRow, true);
        await logAdvisorActivity('edit_record', {
          sheetName: currentSheet,
          rowId: key,
          label: 'Edito registro',
          changedFields: Object.keys(changes),
          changedCount: Object.keys(changes).length,
          contactSnapshot: buildContactSnapshot(formValues)
        });
        saveStatus.textContent = 'Guardado ✓ Pendiente de enviar a principal';
        // Actualizamos la copia local (conservando tracking/CRM) en vez de
        // borrar el cache y re-descargar toda la base.
        upsertRowInLocalState(attachFirebaseMeta(mergedRow, {
          data: mergedRow,
          tracking: getTracking(row),
          crm: getCrm(row),
          source: row.__source || null
        }, row.__rowDocId));
        await loadPage(false, false);
        scheduleFollowUpRefresh(300);
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
      const iso = toInputDateValue(String(raw), type);
      const minAttr = headerName === 'Fecha para contactar' ? ` min="${todayISODate()}"` : '';
      const contactAttr = isContactDateTimeField(headerName) ? ' data-contact-real="1"' : '';
      input = `<input id="${id}" name="${escAttr(h)}" class="in control" type="${type}"
               value="${escAttr(iso)}"${minAttr}${contactAttr} ${!isNew ? 'disabled' : ''} />`;

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
    const originalFollowUpDate = toInputDateValue(originalRow?.['Fecha para contactar'] || '', 'date').slice(0, 10);
    if (!creatingNew && followUpDate && followUpDate === originalFollowUpDate) return;
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

  function renderWhatsAppTemplatesHTML(row) {
    const first = WHATSAPP_TEMPLATES[0];
    return `<section class="whatsapp-panel field">
      <div class="label">WhatsApp con plantilla</div>
      <div class="value whatsapp-template-box">
        <div class="whatsapp-template-controls">
          <select id="waTemplateSelect" class="in">
            ${WHATSAPP_TEMPLATES.map(t => `<option value="${escAttr(t.id)}">${esc(t.label)}</option>`).join('')}
          </select>
          <button id="btnCopyWaTemplate" class="btn" type="button">Copiar</button>
          <a id="btnOpenWaTemplate" class="btn btn-whatsapp-mini" target="_blank" rel="noopener">Abrir WhatsApp</a>
        </div>
        <textarea id="waTemplatePreview" class="in" rows="4">${esc(fillWhatsAppTemplate(first.text, row))}</textarea>
        <span id="waTemplateStatus" class="crm-meta"></span>
      </div>
    </section>`;
  }

  function wireWhatsAppTemplates(row) {
    const select = modalBody.querySelector('#waTemplateSelect');
    const preview = modalBody.querySelector('#waTemplatePreview');
    const copyBtn = modalBody.querySelector('#btnCopyWaTemplate');
    const openBtn = modalBody.querySelector('#btnOpenWaTemplate');
    const statusEl = modalBody.querySelector('#waTemplateStatus');
    if (!select || !preview || !openBtn) return;

    const updateLink = () => {
      const phone = normalizePhone(getField(row, 'phone'));
      const num = phone.startsWith('57') ? phone.slice(2) : phone;
      if (phone.length >= 7) {
        openBtn.href = `https://wa.me/57${num}?text=${encodeURIComponent(preview.value)}`;
        openBtn.classList.remove('disabled');
      } else {
        openBtn.removeAttribute('href');
        openBtn.classList.add('disabled');
      }
    };

    const update = () => {
      const template = WHATSAPP_TEMPLATES.find(t => t.id === select.value) || WHATSAPP_TEMPLATES[0];
      preview.value = fillWhatsAppTemplate(template.text, row);
      updateLink();
      if (statusEl) statusEl.textContent = '';
    };

    select.addEventListener('change', update);
    preview.addEventListener('input', updateLink);
    copyBtn?.addEventListener('click', async () => {
      try {
        await copyText(preview.value);
        if (statusEl) statusEl.textContent = 'Mensaje copiado.';
      } catch (_) {
        if (statusEl) statusEl.textContent = 'No se pudo copiar automaticamente.';
      }
    });
    update();
  }

  function fillWhatsAppTemplate(template, row) {
    const nextDate = getNextContactDate(row);
    const values = {
      nombre: getField(row, 'name') || 'familia',
      asesor: authUser?.displayName || getField(row, 'advisor') || 'Musicala',
      arte: getField(row, 'art') || 'arte',
      modalidad: getField(row, 'modality') || 'la modalidad que prefieras',
      sede: getField(row, ['Sede', 'Ubicacion', 'UbicaciÃ³n']) || 'Musicala',
      fecha: nextDate ? formatYMD(nextDate) : 'la fecha acordada',
      hora: getTemplateTimeValue(row)
    };
    return String(template || '').replace(/\{\{\s*(\w+)\s*\}\}/g, (_, key) => values[key] || '');
  }

  function getTemplateTimeValue(row) {
    const raw = String(getField(row, ['Hora', 'Horario', 'Fecha y hora de contacto', 'Fecha para contactar']) || '');
    const match = raw.match(/(\d{1,2}):(\d{2})/);
    return match ? `${match[1].padStart(2, '0')}:${match[2]}` : 'la hora acordada';
  }

  async function copyText(text) {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return;
    }
    const temp = document.createElement('textarea');
    temp.value = text;
    temp.setAttribute('readonly', '');
    temp.style.position = 'fixed';
    temp.style.opacity = '0';
    document.body.appendChild(temp);
    temp.select();
    document.execCommand('copy');
    document.body.removeChild(temp);
  }

  function renderHistoryHTML(isNew) {
    if (isNew) return '';
    return `<section class="history-panel field">
      <div class="label">Historial</div>
      <div id="historyList" class="value history-list">
        <div class="history-empty">Cargando historial...</div>
      </div>
    </section>`;
  }

  async function loadModalHistory(row) {
    const historyList = modalBody.querySelector('#historyList');
    if (!historyList || !firebaseDb || !currentSheet) return;
    try {
      const doc = sheetDoc(currentSheet);
      const id = row.__rowDocId || rowDocId(row, 0);
      let snap;
      try {
        snap = await doc.collection('rows').doc(id).collection('seguimientos')
          .orderBy('createdAt', 'desc')
          .limit(12)
          .get();
      } catch (_) {
        snap = await doc.collection('rows').doc(id).collection('seguimientos').get();
      }
      const logs = snap.docs.map(d => d.data() || {}).sort((a, b) => Number(b.createdAt || 0) - Number(a.createdAt || 0)).slice(0, 12);
      historyList.innerHTML = logs.length ? logs.map(renderHistoryItem).join('') : '<div class="history-empty">Sin seguimientos registrados todavia.</div>';
    } catch (err) {
      console.warn('No se pudo cargar historial:', err);
      historyList.innerHTML = `<div class="history-empty">No se pudo cargar historial: ${esc(err?.message || err)}</div>`;
    }
  }

  function renderHistoryItem(log) {
    const when = formatTimestamp(log.createdAt);
    const advisor = log.advisorEmail || 'Sin asesor';
    const result = log.resultLabel || log.action || 'Seguimiento';
    const next = log.nextContactAt ? `Proxima: ${formatTimestamp(log.nextContactAt)}` : '';
    const notes = log.notes || log.reason || '';
    return `<article class="history-item">
      <div class="history-item-head">
        <strong>${esc(result)}</strong>
        <span>${esc(when)}</span>
      </div>
      <div class="history-meta">${esc([advisor, next].filter(Boolean).join(' - '))}</div>
      ${notes ? `<p>${esc(notes)}</p>` : ''}
    </article>`;
  }

  function formatTimestamp(value) {
    const parsed = parseFlexibleDate(value);
    if (!parsed) return '';
    const d = typeof value === 'number' ? new Date(value) : parsed;
    return new Intl.DateTimeFormat('es-CO', {
      day:'2-digit',
      month:'short',
      year:'numeric',
      hour:'2-digit',
      minute:'2-digit'
    }).format(d);
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
      const db = window.firebase.firestore(app);
      // Cache local (IndexedDB): al recargar la pagina los datos se leen del equipo
      // y Firestore solo trae lo que cambio. Si el navegador no lo soporta o hay
      // varias pestanas antiguas abiertas, seguimos con red normal sin romper nada.
      try {
        db.enablePersistence({ synchronizeTabs: true }).catch((err) => {
          console.warn('Cache local no disponible, se usa red normal:', err?.code || err);
        });
      } catch (err) {
        console.warn('Cache local no disponible, se usa red normal:', err);
      }
      setFirebaseStatus('Ingresa con Google para ver la base de datos.');
      return db;
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
    [sheetSel, searchBox, btnNew, btnAdvisorStats, btnSyncFirebase, btnSyncSheets].forEach((el) => {
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

  function setupAdvisorAdminUI() {
    if (btnAdvisorStats && isAdminUser(authUser)) {
      btnAdvisorStats.classList.remove('hidden');
      const today = todayISODate();
      if (statsTo && !statsTo.value) statsTo.value = today;
      if (statsFrom && !statsFrom.value) statsFrom.value = formatYMD(addDays(toYMD(new Date()), -6));
    }
  }

  function isAdminUser(user) {
    const email = String(user?.email || '').toLowerCase();
    return ADMIN_EMAILS.includes(email);
  }

  function activityCollection() {
    return firebaseDb ? firebaseDb.collection('advisorActivity') : null;
  }

  function currentAdvisorEmail() {
    return String(authUser?.email || '');
  }

  async function logAdvisorActivity(type, details = {}) {
    if (!firebaseDb || !authUser?.email) return;
    const now = Date.now();
    const email = currentAdvisorEmail();
    const payload = {
      type: String(type || 'event'),
      advisorEmail: email,
      advisorName: String(authUser.displayName || email),
      createdAt: now,
      date: formatYMD(toYMD(new Date(now))),
      sheetName: String(details.sheetName || currentSheet || ''),
      rowId: String(details.rowId || ''),
      label: String(details.label || type || 'Evento'),
      details: sanitizeActivityDetails(details)
    };
    await activityCollection().add(payload).catch((err) => {
      console.warn('No se pudo registrar actividad del asesor:', err);
    });
  }

  function sanitizeActivityDetails(details = {}) {
    const allowed = {};
    Object.entries(details).forEach(([key, value]) => {
      if (['sheetName', 'rowId', 'label'].includes(key)) return;
      if (value == null) allowed[key] = null;
      else if (Array.isArray(value)) allowed[key] = value.map(v => String(v)).slice(0, 80);
      else if (typeof value === 'object') allowed[key] = JSON.parse(JSON.stringify(value));
      else allowed[key] = value;
    });
    return allowed;
  }

  function openAdvisorStats() {
    if (!isAdminUser(authUser)) return;
    advisorStatsPanel?.classList.remove('hidden');
    renderAdvisorStats();
  }

  async function renderAdvisorStats() {
    if (!firebaseDb || !isAdminUser(authUser)) return;
    const from = statsFrom?.value || todayISODate();
    const to = statsTo?.value || from;
    const start = startOfDayMs(from);
    const end = endOfDayMs(to);

    advisorStatsContent.innerHTML = '<div class="alert-empty">Cargando estadisticas...</div>';
    try {
      const snap = await activityCollection()
        .where('createdAt', '>=', start)
        .where('createdAt', '<=', end)
        .get();
      const events = snap.docs.map(d => ({ id: d.id, ...(d.data() || {}) }));
      const byAdvisor = groupAdvisorEvents(events);
      const totals = summarizeAdvisorEvents(events);
      advisorStatsSummary.textContent = `${events.length.toLocaleString()} evento(s) entre ${from} y ${to}.`;
      advisorStatsContent.innerHTML = renderAdvisorStatsHTML(byAdvisor, totals, from, to);
    } catch (err) {
      console.error(err);
      advisorStatsContent.innerHTML = `<div class="alert-empty">No se pudieron cargar las estadisticas: ${esc(err?.message || err)}</div>`;
    }
  }

  function groupAdvisorEvents(events) {
    const grouped = new Map();
    events.forEach((event) => {
      const email = String(event.advisorEmail || 'sin-correo').toLowerCase();
      if (!grouped.has(email)) grouped.set(email, []);
      grouped.get(email).push(event);
    });
    return Array.from(grouped.entries())
      .map(([email, list]) => ({ email, name: list.find(e => e.advisorName)?.advisorName || email, events: list, summary: summarizeAdvisorEvents(list) }))
      .sort((a, b) => b.summary.productiveScore - a.summary.productiveScore);
  }

  function summarizeAdvisorEvents(events) {
    const typeCounts = {};
    const sheetCounts = {};
    const actionCounts = {};
    const editedFields = {};
    const uniqueRows = new Set();
    const days = new Set();

    events.forEach((event) => {
      const type = event.type || 'event';
      typeCounts[type] = (typeCounts[type] || 0) + 1;
      if (event.sheetName) sheetCounts[event.sheetName] = (sheetCounts[event.sheetName] || 0) + 1;
      if (event.rowId) uniqueRows.add(`${event.sheetName || ''}:${event.rowId}`);
      if (event.date) days.add(event.date);
      const details = event.details || {};
      const action = details.resultLabel || details.action;
      if (action) actionCounts[action] = (actionCounts[action] || 0) + 1;
      (details.changedFields || []).forEach((field) => {
        editedFields[field] = (editedFields[field] || 0) + 1;
      });
    });

    const recordsCreated = typeCounts.create_record || 0;
    const recordsEdited = typeCounts.edit_record || 0;
    const trackingActions = typeCounts.tracking_action || 0;
    const syncs = (typeCounts.sync_firebase || 0) + (typeCounts.sync_sheets || 0);
    const changedFieldsTotal = events.reduce((sum, e) => sum + Number(e.details?.changedCount || 0), 0);
    const productiveScore = recordsCreated * 5 + recordsEdited * 3 + trackingActions * 4 + changedFieldsTotal + syncs;

    return {
      totalEvents: events.length,
      recordsCreated,
      recordsEdited,
      trackingActions,
      syncs,
      changedFieldsTotal,
      uniqueRows: uniqueRows.size,
      daysWorked: days.size,
      productiveScore,
      typeCounts,
      sheetCounts,
      actionCounts,
      editedFields
    };
  }

  function renderAdvisorStatsHTML(groups, totals, from, to) {
    if (!groups.length) return '<div class="alert-empty">No hay actividad registrada en este rango.</div>';
    return `
      <div class="stats-overview">
        ${statsMetric('Eventos', totals.totalEvents)}
        ${statsMetric('Registros creados', totals.recordsCreated)}
        ${statsMetric('Ediciones', totals.recordsEdited)}
        ${statsMetric('Gestiones', totals.trackingActions)}
        ${statsMetric('Registros tocados', totals.uniqueRows)}
      </div>
      <div class="advisor-grid">
        ${groups.map(group => renderAdvisorCard(group, from, to)).join('')}
      </div>
    `;
  }

  function renderAdvisorCard(group) {
    const s = group.summary;
    return `
      <article class="advisor-card">
        <div class="advisor-card-head">
          <div>
            <h3>${esc(group.name)}</h3>
            <p>${esc(group.email)}</p>
          </div>
          <strong>${s.productiveScore.toLocaleString()} pts</strong>
        </div>
        <div class="stats-overview compact">
          ${statsMetric('Eventos', s.totalEvents)}
          ${statsMetric('Creados', s.recordsCreated)}
          ${statsMetric('Editados', s.recordsEdited)}
          ${statsMetric('Gestiones', s.trackingActions)}
          ${statsMetric('Campos', s.changedFieldsTotal)}
          ${statsMetric('Dias', s.daysWorked)}
        </div>
        ${statsList('Acciones de seguimiento', s.actionCounts)}
        ${statsList('Hojas trabajadas', s.sheetCounts)}
        ${statsList('Campos mas editados', s.editedFields)}
      </article>
    `;
  }

  function statsMetric(label, value) {
    return `<div class="stats-metric"><span>${esc(label)}</span><strong>${Number(value || 0).toLocaleString()}</strong></div>`;
  }

  function statsList(title, counts) {
    const entries = Object.entries(counts || {}).sort((a, b) => b[1] - a[1]).slice(0, 8);
    if (!entries.length) return '';
    return `
      <div class="stats-list">
        <h4>${esc(title)}</h4>
        ${entries.map(([label, value]) => `<div><span>${esc(label)}</span><strong>${Number(value || 0).toLocaleString()}</strong></div>`).join('')}
      </div>
    `;
  }

  function startOfDayMs(ymd) {
    const d = parseYMD(ymd) || toYMD(new Date());
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0).getTime();
  }

  function endOfDayMs(ymd) {
    const d = parseYMD(ymd) || toYMD(new Date());
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999).getTime();
  }

  function scheduleHideProgress(delay) {
    if (hideProgressTimer) window.clearTimeout(hideProgressTimer);
    hideProgressTimer = window.setTimeout(() => { hideProgressTimer = null; hideProgress(); }, delay);
  }
  function cancelHideProgress() {
    if (hideProgressTimer) { window.clearTimeout(hideProgressTimer); hideProgressTimer = null; }
  }

  function showProgress(title, text, percent = 0) {
    if (!progressPanel) return;
    cancelHideProgress();
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

  // Barra de carga "animada" para operaciones de una sola llamada (sin eventos de avance,
  // como un .get() de Firestore). Sube sola hacia un techo mientras esperamos y se
  // completa con finishFauxProgress() al terminar. Evita que el asesor crea que se colgo.
  function startFauxProgress(title, text, { start = 8, ceiling = 90, stepMs = 350 } = {}) {
    stopFauxProgressTimer();
    showProgress(title, text, start);
    let current = start;
    fauxProgressTimer = window.setInterval(() => {
      // Avanza rapido al principio y mas lento cerca del techo.
      const remaining = ceiling - current;
      const step = Math.max(0.5, remaining * 0.12);
      current = Math.min(ceiling, current + step);
      setProgress(Math.round(current), null);
    }, stepMs);
  }
  function stopFauxProgressTimer() {
    if (fauxProgressTimer) {
      window.clearInterval(fauxProgressTimer);
      fauxProgressTimer = null;
    }
  }
  function finishFauxProgress(text = 'Listo.', hideAfter = 600) {
    stopFauxProgressTimer();
    setProgress(100, text);
    scheduleHideProgress(hideAfter);
  }

  // Cede el control al navegador para que no se congele ("La pagina no responde").
  function yieldToUI() {
    return new Promise((resolve) => window.setTimeout(resolve, 0));
  }

  // Recorre una lista grande en lotes, soltando el hilo entre lotes para mantener
  // la pestana responsiva mientras procesamos miles de registros.
  async function forEachChunked(items, fn, { chunkSize = 400, onProgress = null } = {}) {
    const list = items || [];
    for (let i = 0; i < list.length; i++) {
      fn(list[i], i);
      if ((i + 1) % chunkSize === 0) {
        if (onProgress) onProgress(i + 1, list.length);
        await yieldToUI();
      }
    }
    if (onProgress) onProgress(list.length, list.length);
  }

  function sheetDoc(sheetName) {
    if (!firebaseDb) return null;
    return firebaseDb.collection('sheetCache').doc(slug(sheetName) || 'hoja');
  }

  function rowDocId(row, index) {
    const id = String(row?.ID ?? '').trim();
    return id ? slug(id) : `fila-${index + 1}`;
  }

  function createSyncBatchId(sheetName) {
    return `${slug(sheetName) || 'hoja'}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  }

  function stableStringify(value) {
    if (value == null || typeof value !== 'object') return JSON.stringify(value);
    if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`;
    return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`).join(',')}}`;
  }

  function hashRow(row) {
    const raw = stableStringify(row || {});
    let hash = 2166136261;
    for (let i = 0; i < raw.length; i += 1) {
      hash ^= raw.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }

  function attachFirebaseMeta(row, payload, docId) {
    const out = { ...(payload?.data || row || {}) };
    Object.defineProperty(out, '__tracking', {
      value: payload?.tracking || null,
      enumerable: false,
      configurable: true
    });
    Object.defineProperty(out, '__crm', {
      value: payload?.crm || null,
      enumerable: false,
      configurable: true
    });
    Object.defineProperty(out, '__source', {
      value: payload?.source || null,
      enumerable: false,
      configurable: true
    });
    Object.defineProperty(out, '__rowDocId', {
      value: docId || rowDocId(out, 0),
      enumerable: false,
      configurable: true
    });
    Object.defineProperty(out, '__enrollment', {
      value: payload?.enrollment || null,
      enumerable: false,
      configurable: true,
      writable: true
    });
    return out;
  }

  function getTracking(row) {
    return row?.__tracking && typeof row.__tracking === 'object' ? row.__tracking : {};
  }

  function getCrm(row) {
    return row?.__crm && typeof row.__crm === 'object' ? row.__crm : {};
  }

  function cloneTracking(tracking) {
    return tracking && typeof tracking === 'object' ? { ...tracking } : {};
  }

  async function loadSheetNamesFromFirebase() {
    if (!firebaseDb) return [];
    try {
      const snap = await firebaseDb.collection('sheetCache').get();
      const names = snap.docs
        .map(docSnap => {
          const data = docSnap.data() || {};
          return String(data.sheetName || docSnap.id || '').trim();
        })
        .filter(Boolean);
      return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b, 'es', { sensitivity:'base' }));
    } catch (err) {
      console.warn('No se pudieron cargar hojas desde Firebase:', err);
      setFirebaseStatus('No se pudieron cargar las hojas guardadas en Firebase.');
      return [];
    }
  }

  async function loadSheetFromFirebase(sheetName) {
    if (!firebaseDb) return null;

    try {
      const doc = sheetDoc(sheetName);
      const metaSnap = await doc.get();
      if (!metaSnap.exists) return null;

      let rowSnap;
      try {
        rowSnap = await doc.collection('rows')
          .where('active', '==', true)
          .orderBy('source.rowIndex')
          .get();
      } catch (indexedErr) {
        console.warn('No se pudo leer ordenado desde Firebase, usando lectura compatible:', indexedErr);
        rowSnap = await doc.collection('rows').get();
      }
      if (rowSnap.empty) return null;

      const meta = metaSnap.data() || {};
      const rows = rowSnap.docs
        .map(d => ({ id: d.id, payload: d.data() || {} }))
        .filter(item => item.payload.active !== false)
        .sort((a, b) => {
          const ai = Number(a.payload?.source?.rowIndex);
          const bi = Number(b.payload?.source?.rowIndex);
          if (Number.isFinite(ai) && Number.isFinite(bi) && ai !== bi) return ai - bi;
          if (Number.isFinite(ai)) return -1;
          if (Number.isFinite(bi)) return 1;
          return a.id.localeCompare(b.id);
        })
        .map(item => attachFirebaseMeta(null, item.payload, item.id));
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

  async function loadSheetPageFromFirebase(sheetName, pageOffset = 0, pageLimit = 200) {
    if (!firebaseDb) return null;

    try {
      const doc = sheetDoc(sheetName);
      const metaSnap = await doc.get();
      if (!metaSnap.exists) return null;
      const meta = metaSnap.data() || {};
      const readLimit = Math.max(pageLimit, pageOffset + pageLimit);

      let rowSnap;
      try {
        rowSnap = await doc.collection('rows')
          .where('active', '==', true)
          .orderBy('source.rowIndex')
          .limit(readLimit)
          .get();
      } catch (indexedErr) {
        console.warn('No se pudo leer pagina ordenada desde Firebase, usando lectura compatible:', indexedErr);
        rowSnap = await doc.collection('rows').limit(readLimit).get();
      }
      if (rowSnap.empty) return null;

      const rows = rowSnap.docs
        .map(d => ({ id: d.id, payload: d.data() || {} }))
        .filter(item => item.payload.active !== false)
        .sort((a, b) => {
          const ai = Number(a.payload?.source?.rowIndex);
          const bi = Number(b.payload?.source?.rowIndex);
          if (Number.isFinite(ai) && Number.isFinite(bi) && ai !== bi) return ai - bi;
          if (Number.isFinite(ai)) return -1;
          if (Number.isFinite(bi)) return 1;
          return a.id.localeCompare(b.id);
        })
        .slice(pageOffset, pageOffset + pageLimit)
        .map(item => attachFirebaseMeta(null, item.payload, item.id));

      const headers = Array.isArray(meta.headers) && meta.headers.length
        ? meta.headers
        : Array.from(rows.reduce((set, row) => {
            Object.keys(row || {}).forEach(k => set.add(k));
            return set;
          }, new Set()));

      const totalRows = Number(meta.total || 0) || Math.max(rows.length, pageOffset + rows.length);
      setFirebaseStatus(`Base lista: ${totalRows.toLocaleString()} registros disponibles.`);
      return { headers, rows, total: totalRows, ts: meta.updatedAt || Date.now(), source: 'firebase-page' };
    } catch (err) {
      console.warn('No se pudo leer la pagina rapida, se usara respaldo:', err);
      setFirebaseStatus('Cargando informaciÃ³n actualizada...');
      return null;
    }
  }

  async function saveSheetToFirebase(sheetName, payload) {
    if (!firebaseDb || !payload?.rows?.length) return;

    const doc = sheetDoc(sheetName);
    const batchSize = 400;
    const rows = payload.rows || [];
    const now = Date.now();
    const syncBatchId = createSyncBatchId(sheetName);
    const seenIds = new Set();

    await doc.set({
      sheetName,
      headers: payload.headers || [],
      total: rows.length,
      updatedAt: now,
      lastSyncBatchId: syncBatchId,
      source: 'sheets'
    }, { merge: true });

    for (let start = 0; start < rows.length; start += batchSize) {
      const batch = firebaseDb.batch();
      rows.slice(start, start + batchSize).forEach((row, i) => {
        const rowIndex = start + i;
        const rowId = rowDocId(row, rowIndex);
        seenIds.add(rowId);
        batch.set(doc.collection('rows').doc(rowId), {
          data: row,
          crm: normalizeCrm(row),
          source: {
            sheetName,
            sheetSlug: slug(sheetName),
            rowIndex,
            sheetRowNumber: rowIndex + 2,
            updatedFromSheetAt: now,
            syncBatchId
          },
          active: true,
          lastSeenAt: now,
          lastSheetHash: hashRow(row),
          syncError: null,
          updatedAt: now,
          version: window.firebase.firestore.FieldValue.increment(1)
        }, { merge: true });
      });
      await batch.commit();
    }

    await archiveRowsMissingFromSync(doc, syncBatchId, seenIds);

    setFirebaseStatus(`Datos actualizados: ${rows.length.toLocaleString()} registros disponibles.`);
  }

  async function archiveRowsMissingFromSync(doc, syncBatchId, seenIds) {
    const snap = await doc.collection('rows').get();
    const stale = snap.docs.filter((rowDoc) => {
      const data = rowDoc.data() || {};
      if (seenIds.has(rowDoc.id)) return false;
      if (data.active === false) return false;
      if (data.pendingSheetSync === true) return false;
      return data?.source?.syncBatchId !== syncBatchId;
    });

    const batchSize = 400;
    for (let start = 0; start < stale.length; start += batchSize) {
      const batch = firebaseDb.batch();
      stale.slice(start, start + batchSize).forEach((rowDoc) => {
        batch.set(rowDoc.ref, {
          active: false,
          archivedBySync: true,
          archivedAt: Date.now(),
          updatedAt: Date.now()
        }, { merge: true });
      });
      await batch.commit();
    }
  }

  async function saveRowToFirebase(sheetName, row, pendingSheetSync = false) {
    if (!firebaseDb) throw new Error('No se pudo guardar el cambio rapido.');
    const doc = sheetDoc(sheetName);
    const id = rowDocId(row, 0);
    if (!id) throw new Error('El registro no tiene ID.');
    const now = Date.now();

    await doc.collection('rows').doc(id).set({
      data: row,
      crm: normalizeCrm(row),
      source: {
        ...(row.__source || {}),
        sheetName,
        sheetSlug: slug(sheetName)
      },
      active: true,
      updatedAt: now,
      pendingSheetSync,
      pendingReason: pendingSheetSync ? 'edit' : '',
      syncedAt: pendingSheetSync ? null : now,
      syncError: null,
      version: window.firebase.firestore.FieldValue.increment(1)
    }, { merge: true });

    await doc.set({
      sheetName,
      updatedAt: now
    }, { merge: true });
  }

  async function updateRowInSheets(sheetName, row) {
    if (!API_BASE) throw new Error('El respaldo de Google Sheets no está configurado (window.API_BASE).');
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
      return 0;
    }

    if (showMessages) status(`Guardando ${snap.size.toLocaleString()} cambio(s) en la base principal...`);

    let done = 0;
    let failed = 0;
    for (const rowDoc of snap.docs) {
      const payload = rowDoc.data() || {};
      const row = payload.data || {};
      setProgress(Math.round((done / snap.size) * 100), `Guardando cambios en la base principal (${done + 1} de ${snap.size})...`);
      try {
        await updateRowInSheets(currentSheet, row);
        await rowDoc.ref.set({
          pendingSheetSync: false,
          pendingReason: '',
          syncedAt: Date.now(),
          crm: normalizeCrm(row),
          syncError: null,
          lastSheetHash: hashRow(row),
          version: window.firebase.firestore.FieldValue.increment(1)
        }, { merge: true });
      } catch (err) {
        failed += 1;
        await rowDoc.ref.set({
          pendingSheetSync: true,
          syncError: String(err?.message || err),
          updatedAt: Date.now()
        }, { merge: true });
        console.warn('No se pudo guardar fila pendiente en Sheets:', err);
      }
      done += 1;
    }

    allRowsCache.delete(currentSheet);
    if (showMessages) {
      const ok = done - failed;
      status(failed ? `${ok.toLocaleString()} cambio(s) guardados; ${failed.toLocaleString()} quedaron con error.` : `${ok.toLocaleString()} cambio(s) guardados en la base principal.`);
    }
    return done - failed;
  }

  function scheduleFollowUpRefresh(delay = 2500) {
    if (!alertsEl && !todayViewEl) return;
    const token = ++followUpRefreshToken;
    if (followUpRefreshTimer) window.clearTimeout(followUpRefreshTimer);
    followUpRefreshTimer = window.setTimeout(async () => {
      if (token !== followUpRefreshToken) return;
      const run = async () => {
        if (token !== followUpRefreshToken) return;
        try {
          await computeAndRenderAlerts();
        } catch (err) {
          console.warn('No se pudieron calcular Hoy/alertas:', err);
        }
      };
      if ('requestIdleCallback' in window) window.requestIdleCallback(run, { timeout: 5000 });
      else run();
    }, delay);
  }

  function scheduleDailySyncIfNeeded() {
    if (dailySyncScheduled) return;
    dailySyncScheduled = true;
    window.setTimeout(async () => {
      try {
        await runDailySyncIfNeeded();
      } catch (err) {
        console.warn('No se pudo completar la sincronizacion diaria en segundo plano:', err);
      }
    }, 8000);
  }

  function renderFollowUpPlaceholder() {
    if (todayViewEl) {
      todayViewEl.classList.remove('hidden');
      todayViewEl.classList.add('today-collapsed');
      todayViewEl.innerHTML = `
        <div class="today-head">
          <div>
            <span class="alert-eyebrow">Hoy</span>
            <h2>Hoy en Musicala</h2>
            <p>Cargando la base primero. Hoy y el seguimiento guiado se preparan automaticamente en segundo plano.</p>
          </div>
          <div class="today-head-actions">
            <button id="btnLoadTodayView" class="btn today-toggle" type="button">Preparando...</button>
          </div>
        </div>
      `;
      todayViewEl.querySelector('#btnLoadTodayView')?.addEventListener('click', async () => {
        const btn = todayViewEl.querySelector('#btnLoadTodayView');
        if (btn) {
          btn.disabled = true;
          btn.textContent = 'Cargando...';
        }
        await computeAndRenderAlerts();
      });
    }
    if (alertsEl) {
      alertsEl.classList.add('hidden');
      alertsEl.innerHTML = '';
    }
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
    const firebaseFresh = await loadSheetFromFirebase(currentSheet);
    allRowsCache.set(currentSheet, firebaseFresh || { ...fresh, source: 'firebase' });
    if (showMessages) setProgress(85, 'Preparando vista rapida...');
    renderFollowUpPlaceholder();
    scheduleFollowUpRefresh(250);
    if (showMessages) {
      setProgress(100, 'Datos actualizados correctamente.');
      status(`Datos actualizados para "${currentSheet}".`);
      window.setTimeout(hideProgress, 900);
    }
  }

  /* =============================
     6.5) Conciliación con Lista de estudiantes
     (proyecto Firebase separado: estudiantes-musicala, SOLO lectura)
     ============================= */
  let listaAppInstance = null;
  let lastReconcileMatches = null;

  function getListaApp() {
    const cfg = window.LISTA_FIREBASE_CONFIG || {};
    if (!cfg.apiKey || !cfg.projectId || !window.firebase?.initializeApp) {
      throw new Error('Falta la configuración de Lista de estudiantes (window.LISTA_FIREBASE_CONFIG).');
    }
    if (!listaAppInstance) {
      listaAppInstance = (window.firebase.apps || []).find((a) => a.name === 'lista')
        || window.firebase.initializeApp(cfg, 'lista');
    }
    return listaAppInstance;
  }

  async function ensureListaAuth(app) {
    const auth = app.auth();
    if (auth.currentUser) return auth.currentUser;
    const provider = new window.firebase.auth.GoogleAuthProvider();
    const hint = window.firebase.auth().currentUser?.email;
    if (hint) provider.setCustomParameters({ login_hint: hint });
    const cred = await auth.signInWithPopup(provider);
    return cred.user;
  }

  function digitsKey(v) {
    const d = String(v ?? '').replace(/\D+/g, '');
    return d.length > 10 ? d.slice(-10) : d;
  }

  function nameMatchKey(v) {
    return String(v ?? '')
      .toLowerCase()
      .normalize('NFD').replace(/\p{Diacritic}/gu, '')
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function emailMatchKey(v) {
    return String(v ?? '').trim().toLowerCase();
  }

  function rowValueByAliases(row, aliases) {
    for (const alias of aliases) {
      const val = row?.[alias];
      if (val != null && String(val).trim() !== '') return val;
    }
    // Búsqueda flexible por si el encabezado tiene tildes/codificación distinta
    const wanted = aliases.map((a) => nameMatchKey(a));
    for (const key of Object.keys(row || {})) {
      if (wanted.includes(nameMatchKey(key))) {
        const val = row[key];
        if (val != null && String(val).trim() !== '') return val;
      }
    }
    return '';
  }

  async function loadEstudiantesFromLista() {
    const app = getListaApp();
    await ensureListaAuth(app);
    const db = window.firebase.firestore(app);
    const snap = await db.collection('estudiantes').get();
    return snap.docs.map((d) => ({ id: d.id, ...(d.data() || {}) }));
  }

  function buildStudentIndexes(students) {
    const byPhone = new Map();
    const byEmail = new Map();
    const byName = new Map();
    const add = (map, key, student) => { if (key && !map.has(key)) map.set(key, student); };
    students.forEach((s) => {
      [s.mobile, s.guardianMobile, s.phone, s.guardianPhone].forEach((v) => {
        const key = digitsKey(v);
        if (key.length >= 7) add(byPhone, key, s);
      });
      add(byEmail, emailMatchKey(s.studentEmail), s);
      add(byName, nameMatchKey(s.studentName), s);
    });
    return { byPhone, byEmail, byName };
  }

  function matchRowToStudent(row, indexes) {
    const phone = digitsKey(rowValueByAliases(row, FIELD_ALIASES.phone));
    if (phone.length >= 7 && indexes.byPhone.has(phone)) {
      return { student: indexes.byPhone.get(phone), matchedBy: 'celular' };
    }
    const email = emailMatchKey(rowValueByAliases(row, FIELD_ALIASES.email));
    if (email && indexes.byEmail.has(email)) {
      return { student: indexes.byEmail.get(email), matchedBy: 'correo' };
    }
    const name = nameMatchKey(rowValueByAliases(row, FIELD_ALIASES.name));
    if (name && name.split(' ').length >= 2 && indexes.byName.has(name)) {
      return { student: indexes.byName.get(name), matchedBy: 'nombre' };
    }
    return null;
  }

  async function runReconciliation() {
    if (reconcileSummary) reconcileSummary.textContent = 'Leyendo Lista de estudiantes...';
    if (reconcileContent) reconcileContent.innerHTML = '';
    btnSaveReconcile?.classList.add('hidden');

    const [students, cache] = await Promise.all([
      loadEstudiantesFromLista(),
      loadAllRowsForSheet(currentSheet)
    ]);
    const rows = cache?.rows || [];
    if (!students.length) {
      if (reconcileSummary) reconcileSummary.textContent = 'La Lista de estudiantes no devolvió registros.';
      return;
    }

    const indexes = buildStudentIndexes(students);
    const matches = [];
    rows.forEach((row) => {
      const hit = matchRowToStudent(row, indexes);
      if (hit) matches.push({ row, docId: row.__rowDocId, student: hit.student, matchedBy: hit.matchedBy });
    });
    lastReconcileMatches = matches;

    const statusCount = {};
    matches.forEach((m) => {
      const st = String(m.student.status || 'Sin estado').trim() || 'Sin estado';
      statusCount[st] = (statusCount[st] || 0) + 1;
    });

    if (reconcileSummary) {
      reconcileSummary.textContent =
        `${students.length.toLocaleString()} estudiantes en la Lista · ` +
        `${matches.length.toLocaleString()} coincidencias en "${currentSheet}" · ` +
        Object.entries(statusCount).map(([k, v]) => `${k}: ${v.toLocaleString()}`).join(' · ');
    }

    if (reconcileContent) {
      if (!matches.length) {
        reconcileContent.innerHTML = '<p>No se encontraron coincidencias por celular, correo o nombre en esta hoja.</p>';
      } else {
        const rowsHtml = matches.map((m) => `
          <tr>
            <td>${esc(rowValueByAliases(m.row, FIELD_ALIASES.name) || '(sin nombre)')}</td>
            <td>${esc(m.student.studentName || '')}</td>
            <td>${renderEnrollmentBadge({ status: m.student.status })}</td>
            <td>${esc(m.matchedBy)}</td>
          </tr>`).join('');
        reconcileContent.innerHTML = `
          <div class="table-wrap">
            <table class="data-table">
              <thead><tr><th>En base comercial</th><th>En Lista de estudiantes</th><th>Estado</th><th>Coincidencia por</th></tr></thead>
              <tbody>${rowsHtml}</tbody>
            </table>
          </div>`;
      }
    }
    if (matches.length) btnSaveReconcile?.classList.remove('hidden');
  }

  async function saveReconciliationToFirebase() {
    if (!firebaseDb) throw new Error('Sin conexión con la base.');
    const matches = lastReconcileMatches || [];
    if (!matches.length) throw new Error('Primero ejecuta la conciliación.');

    const doc = sheetDoc(currentSheet);
    const now = Date.now();
    const batchSize = 400;
    for (let start = 0; start < matches.length; start += batchSize) {
      const batch = firebaseDb.batch();
      matches.slice(start, start + batchSize).forEach((m) => {
        if (!m.docId) return;
        batch.set(doc.collection('rows').doc(m.docId), {
          enrollment: {
            status: String(m.student.status || ''),
            studentId: String(m.student.id || ''),
            studentName: String(m.student.studentName || ''),
            matchedBy: m.matchedBy,
            checkedAt: now
          },
          updatedAt: now
        }, { merge: true });
      });
      await batch.commit();
    }

    // Reflejar en la vista sin re-descargar la base
    matches.forEach((m) => {
      try { m.row.__enrollment = { status: m.student.status || '', matchedBy: m.matchedBy, checkedAt: now }; } catch (_) {}
    });
    renderTable(currentHeaders, currentRows);
    if (reconcileSummary) {
      reconcileSummary.textContent = `Conciliación guardada: ${matches.length.toLocaleString()} registros marcados con su estado de inscripción.`;
    }
    await logAdvisorActivity('reconcile_lista', {
      sheetName: currentSheet,
      label: 'Concilió con Lista de estudiantes',
      count: matches.length
    }).catch(() => {});
  }

  function renderEnrollmentBadge(enrollment) {
    if (!enrollment || (!enrollment.status && !enrollment.matchedBy)) return '';
    const st = String(enrollment.status || 'Inscrito').trim() || 'Inscrito';
    const cls = /activo/i.test(st) ? 'badge music' : /inactivo|retirado/i.test(st) ? 'badge theatre' : 'badge arts';
    return `<span class="${cls}" title="Según Lista de estudiantes">${esc(st)}</span>`;
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

  function escapeHtml(s) {
    return String(s ?? '').replace(/[&<>\"']/g, (c) =>
      ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])
    );
  }
  function escapeAttr(s) {
    return escapeHtml(s);
  }
  function esc(s) { return escapeHtml(s); }
  function escAttr(s){ return escapeAttr(s); }

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
    if (key === '__crmStage') {
      return String(normalizeCrm(a).stageLabel).localeCompare(String(normalizeCrm(b).stageLabel), 'es', { sensitivity: 'base' });
    }
    if (key === '__crmTemperature') {
      return Number(normalizeCrm(a).leadScore || 0) - Number(normalizeCrm(b).leadScore || 0);
    }
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
  async function computeAndRenderAlertsLegacy() {
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

    renderAlertsUILegacy(
      urgent.slice(0,8).map(x=>x.row),
      soon.slice(0,4).map(x=>x.row),
      urgent.length,
      soon.length
    );
  }

  function renderAlertsUILegacy(urgentTop, soonTop, uCount, sCount){
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

  function classifyUrgencyLegacy(row){
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

  function getLastContactDateLegacy(row) {
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

  async function computeAndRenderAlertsBucketsLegacy() {
    if (!alertsEl) return;

    let cache = allRowsCache.get(currentSheet);
    if (!cache) {
      cache = await loadAllRowsForSheet(currentSheet);
      allRowsCache.set(currentSheet, cache);
    }

    const counts = emptyFollowUpCounts();
    const items = [];

    (cache.rows || []).forEach((row) => {
      const urgency = classifyUrgency(row);
      counts[urgency.bucket] = (counts[urgency.bucket] || 0) + 1;
      if (shouldShowInAlerts(urgency)) items.push({ row, urgency });
    });

    items.sort(compareFollowUpItems);
    console.info('[seguimiento] buckets', counts);
    renderAlertsUIBucketsLegacy(items, counts);
  }

  async function saveTrackingToFirebase(sheetName, row, action, extra = {}) {
    if (!firebaseDb) throw new Error('No se pudo guardar el seguimiento en Firebase.');
    const doc = sheetDoc(sheetName);
    const id = rowDocId(row, 0);
    if (!id) throw new Error('El registro no tiene ID.');

    const rowRef = doc.collection('rows').doc(id);
    const previousTracking = cloneTracking(getTracking(row));
    const updatedRow = buildManagedRow(row, action, extra);
    const tracking = buildTrackingUpdate(row, action, extra);

    await rowRef.set({
      data: { ...updatedRow },
      tracking,
      crm: normalizeCrm(updatedRow),
      source: {
        ...(row.__source || {}),
        sheetName,
        sheetSlug: slug(sheetName)
      },
      active: true,
      pendingSheetSync: true,
      pendingReason: `tracking:${tracking.lastAction || action}`,
      syncError: null,
      updatedAt: Date.now(),
      version: window.firebase.firestore.FieldValue.increment(1)
    }, { merge: true });

    const now = Date.now();

    await rowRef.collection('seguimientos').add({
      action: tracking.lastAction || action,
      resultLabel: tracking.resultLabel || action,
      advisorEmail: tracking.advisorEmail || '',
      createdAt: now,
      previousTracking,
      newTracking: tracking,
      contactSnapshot: buildContactSnapshot(updatedRow),
      notes: extra.notes || null,
      reason: extra.reason || null,
      nextContactAt: tracking.nextContactAt || null
    });

    await logAdvisorActivity('tracking_action', {
      sheetName,
      rowId: id,
      action: tracking.lastAction || action,
      resultLabel: tracking.resultLabel || action,
      label: `Gestion: ${tracking.resultLabel || action}`,
      contactSnapshot: buildContactSnapshot(updatedRow),
      nextContactAt: tracking.nextContactAt || null,
      reason: extra.reason || null
    });

    await doc.set({
      sheetName,
      updatedAt: now
    }, { merge: true });

    return updatedRow;
  }

  function renderAlertsUIBucketsLegacy(items, counts){
    if (!alertsEl) return;

    const urgentRealCount = realUrgentCount(counts);
    const soonCount = counts.soon || 0;
    if (!items.length) {
      alertsEl.classList.add('hidden');
      alertsEl.innerHTML = '';
      return;
    }

    const filterDefs = [
      ['all', 'Todos', items.length],
      ['real', 'Urgentes reales', urgentRealCount],
      ['today', 'Hoy', counts.today || 0],
      ['soon', 'Proximos', soonCount],
      ['old', 'Antiguos por revisar', counts.old_review || 0]
    ];
    const filtered = filterAlertItems(items);
    const maxItems = alertFilter === 'old' ? 12 : 18;

    const makeItem = ({ row, urgency }) => {
      const name = String(getField(row, ['Nombre']) || '').trim() || '(Sin nombre)';
      const canal = getField(row, ['Canal de comunicacion', 'Canal de comunicaciÃ³n', 'Canal']) || 'Sin canal';
      const asesor = getField(row, ['Asesor', 'Responsable']);
      const id = String(getField(row, ['ID']) || '').trim();
      const meta = [canal, asesor ? `Asesor: ${asesor}` : '', urgency.reason].filter(Boolean).join(' - ');
      const dateTxt = followUpDateText(urgency);

      return `
        <div class="alert-item alert-${escAttr(urgency.bucket)}">
          <div class="alert-person">
            <strong>${esc(name)}</strong>
            <span class="meta">${esc(meta)}</span>
            ${dateTxt ? `<span class="meta">${esc(dateTxt)}</span>` : ''}
          </div>
          <span class="mini-priority">${esc(urgency.label || 'Pendiente')}</span>
          <button class="btn alert-open" data-openid="${escAttr(id)}">Abrir</button>
        </div>
      `;
    };

    alertsEl.innerHTML = `
      <div class="alert-card">
        <div class="alert-title">
          <div>
            <span class="alert-eyebrow">Seguimiento</span>
            <h2>Contactos que necesitan atenciÃ³n</h2>
          </div>
          <div class="pills">
            <span class="pill urgent">${urgentRealCount.toLocaleString()} urgentes reales</span>
            <span class="pill soon">${soonCount.toLocaleString()} proximos</span>
            <span class="pill review">${(counts.old_review || 0).toLocaleString()} antiguos</span>
          </div>
        </div>
        <div class="alert-summary">
          ${summaryPill('Vencidos', counts.overdue)}
          ${summaryPill('Hoy', counts.today)}
          ${summaryPill('Alta sin fecha', counts.high_no_date)}
          ${summaryPill('Nuevos', counts.fresh_uncontacted)}
          ${summaryPill('Seguimiento', counts.followup)}
          ${summaryPill('Antiguos', counts.old_review)}
        </div>
        <div class="alert-filters">
          ${filterDefs.map(([key, label, count]) => `
            <button type="button" class="alert-filter ${alertFilter === key ? 'active' : ''}" data-alert-filter="${escAttr(key)}">
              ${esc(label)} <span>${Number(count || 0).toLocaleString()}</span>
            </button>
          `).join('')}
        </div>
        <div class="alert-list">
          ${filtered.slice(0, maxItems).map(makeItem).join('')}
        </div>
      </div>
    `;
    alertsEl.classList.remove('hidden');

    alertsEl.querySelectorAll('[data-alert-filter]').forEach(btn => {
      btn.addEventListener('click', () => {
        alertFilter = btn.getAttribute('data-alert-filter') || 'all';
        renderAlertsUIBucketsLegacy(items, counts);
      });
    });

    alertsEl.querySelectorAll('[data-openid]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-openid') || '';
        const row = await getRowById(id);
        if (row) {
          openModal(row);
          setTimeout(()=>document.getElementById('btnEdit')?.click(), 0);
        } else {
          alert('No se encontrÃ³ la fila seleccionada.');
        }
      });
    });
  }

  function classifyUrgency(row){
    if (isExcludedFromFollowUp(row)) {
      return { type:'excluded', level:'excluded', bucket:'excluded', label:'Excluido', reason:'Estado no requiere seguimiento', whenSort:null };
    }

    const priority = normalizeText(getField(row, ['Prioridad']));
    const nextContact = getNextContactDate(row);
    const lastContact = getLastContactDate(row);
    const created = getCreatedDate(row);
    const today = toYMD(new Date());
    const in3 = addDays(today, 3);

    if (nextContact) {
      if (nextContact < today) return buildUrgency('urgent', 'urgent', 'overdue', 'Vencido', 'Fecha de seguimiento vencida', nextContact, 'Para contactar');
      if (nextContact.getTime() === today.getTime()) return buildUrgency('today', 'today', 'today', 'Hoy', 'Seguimiento programado para hoy', nextContact, 'Para contactar');
      if (nextContact <= in3) return buildUrgency('soon', 'soon', 'soon', 'Proximo', 'Seguimiento proximo', nextContact, 'Para contactar');
      return buildUrgency('ok', 'ok', 'scheduled', 'Programado', 'Seguimiento programado a futuro', nextContact, 'Para contactar');
    }

    if (priority === 'alta') {
      return { type:'urgent', level:'urgent', bucket:'high_no_date', label:'Alta sin fecha', reason:'Prioridad alta sin fecha de seguimiento', when:'', whenSort:today };
    }

    if (lastContact) {
      const days = daysBetween(lastContact, today);
      if (days >= 30) return buildUrgency('review', 'review', 'old_review', 'Revisar', 'Ultimo contacto hace 30+ dias', lastContact, 'Ultimo contacto', `${days} dias`);
      if (days >= 15) return buildUrgency('urgent', 'urgent', 'followup', 'Seguimiento', 'Ultimo contacto hace 15+ dias', lastContact, 'Ultimo contacto', `${days} dias`);
      if (days >= 7) return buildUrgency('soon', 'soon', 'soon', 'Proximo', 'Ultimo contacto hace 7+ dias', lastContact, 'Ultimo contacto', `${days} dias`);
      return buildUrgency('ok', 'ok', 'ok', 'Al dia', 'Contacto reciente', lastContact, 'Ultimo contacto');
    }

    if (created) {
      const age = daysBetween(created, today);
      if (age <= 30) return buildUrgency('urgent', 'urgent', 'fresh_uncontacted', 'Nuevo sin contactar', 'Contacto reciente sin primer contacto', created, 'Creado', `${age} dias`);
      return buildUrgency('review', 'review', 'old_review', 'Antiguo por revisar', 'Contacto antiguo sin primer contacto', created, 'Creado', `${age} dias`);
    }

    return { type:'review', level:'review', bucket:'old_review', label:'Sin fecha por revisar', reason:'Sin contacto registrado ni fecha de creacion', when:'', whenSort:null };
  }

  function buildUrgency(type, level, bucket, label, reason, date, dateLabel, whenText){
    return {
      type,
      level,
      bucket,
      label,
      reason,
      when: whenText || formatIf(date),
      whenSort: date || null,
      relevantDate: date || null,
      relevantDateLabel: dateLabel || ''
    };
  }

  function emptyFollowUpCounts(){
    return {
      overdue:0, today:0, high_no_date:0, fresh_uncontacted:0, followup:0,
      soon:0, old_review:0, scheduled:0, ok:0, excluded:0
    };
  }

  function realUrgentCount(counts){
    return ['overdue', 'today', 'high_no_date', 'fresh_uncontacted', 'followup']
      .reduce((sum, key) => sum + Number(counts[key] || 0), 0);
  }

  function summaryPill(label, count){
    return `<span>${esc(label)} <strong>${Number(count || 0).toLocaleString()}</strong></span>`;
  }

  function shouldShowInAlerts(urgency){
    return ['overdue', 'today', 'high_no_date', 'fresh_uncontacted', 'followup', 'soon', 'old_review'].includes(urgency.bucket);
  }

  function filterAlertItems(items){
    if (alertFilter === 'real') return items.filter(x => ['overdue', 'today', 'high_no_date', 'fresh_uncontacted', 'followup'].includes(x.urgency.bucket));
    if (alertFilter === 'today') return items.filter(x => x.urgency.bucket === 'today');
    if (alertFilter === 'soon') return items.filter(x => x.urgency.bucket === 'soon');
    if (alertFilter === 'old') return items.filter(x => x.urgency.bucket === 'old_review');
    return items;
  }

  function compareFollowUpItems(a, b){
    const order = { overdue:1, today:2, high_no_date:3, fresh_uncontacted:4, followup:5, soon:6, old_review:7 };
    const ao = order[a.urgency.bucket] || 99;
    const bo = order[b.urgency.bucket] || 99;
    if (ao !== bo) return ao - bo;

    const ad = a.urgency.whenSort;
    const bd = b.urgency.whenSort;
    if (ad && bd && ad.getTime() !== bd.getTime()) {
      if (a.urgency.bucket === 'fresh_uncontacted' || a.urgency.bucket === 'old_review') return bd - ad;
      return ad - bd;
    }
    if (ad && !bd) return -1;
    if (!ad && bd) return 1;
    return String(getField(a.row, ['Nombre']) || '').localeCompare(String(getField(b.row, ['Nombre']) || ''), 'es', { sensitivity:'base' });
  }

  function followUpDateText(urgency){
    if (!urgency?.relevantDate || !urgency?.relevantDateLabel) return '';
    return `${urgency.relevantDateLabel}: ${formatYMD(urgency.relevantDate)}`;
  }

  function getLastContactDate(row) {
    const tracking = getTracking(row);
    const trackedDate = parseFlexibleDate(tracking.lastContactAt);
    if (trackedDate) return trackedDate;
    return firstDateFrom(row, [
      'Ultimo contacto', 'Ultimo Contacto', 'Ãšltimo contacto', 'Ãšltimo Contacto',
      'Fecha y hora de contacto', 'Fecha de contacto', 'Fecha Contacto'
    ]);
  }

  function getNextContactDate(row) {
    const tracking = getTracking(row);
    const trackedDate = parseFlexibleDate(tracking.nextContactAt);
    if (trackedDate) return trackedDate;
    return firstDateFrom(row, ['Fecha para contactar', 'Siguiente Contacto (calc)']);
  }

  function getCreatedDate(row) {
    return firstDateFrom(row, [
      'Fecha de creacion', 'Fecha creacion', 'Creado', 'Fecha ingreso', 'Fecha de ingreso',
      'Fecha registro', 'Fecha de registro', 'Fecha', 'Created At', 'created_at'
    ]);
  }

  function firstDateFrom(row, names){
    for (const name of names) {
      const parsed = parseFlexibleDate(getField(row, [name]));
      if (parsed) return parsed;
    }
    return null;
  }

  function fieldNames(fieldOrNames) {
    if (Array.isArray(fieldOrNames)) return fieldOrNames;
    const key = String(fieldOrNames || '').trim();
    return FIELD_ALIASES[key] || (key ? [key] : []);
  }

  function getField(row, possibleNames){
    if (!row) return '';
    possibleNames = fieldNames(possibleNames);
    const direct = possibleNames.find(name => Object.prototype.hasOwnProperty.call(row, name));
    if (direct) return row[direct];

    const wanted = possibleNames.map(normalizeText);
    const key = Object.keys(row).find(k => wanted.includes(normalizeText(k)));
    if (key) return row[key];

    for (const name of possibleNames) {
      const found = findHeaderIncludes(row, normalizeText(name).split(/\s+/).filter(Boolean));
      if (found) return row[found];
    }
    return '';
  }

  function setField(row, fieldOrNames, value){
    if (!row) return '';
    const names = fieldNames(fieldOrNames);
    let key = findExistingKey(row, names);
    if (!key) key = names[0] || '';
    if (key) row[key] = value;
    return key;
  }

  function hasField(row, fieldOrNames){
    if (!row) return false;
    return Boolean(findExistingKey(row, fieldNames(fieldOrNames)));
  }

  function normalizeCrm(row) {
    const existing = getCrm(row);
    const stage = existing.stage || inferCrmStage(row);
    const leadScore = Number.isFinite(Number(existing.leadScore)) ? Number(existing.leadScore) : calculateLeadScore(row, stage);
    const temperature = existing.temperature || scoreToTemperature(leadScore);
    const tracking = getTracking(row);
    const nextContact = parseFlexibleDate(tracking.nextContactAt) || getNextContactDate(row);
    const lastActionAt = parseFlexibleDate(tracking.lastManagedAt || tracking.updatedAt);
    const crm = {
      stage,
      temperature,
      leadScore,
      nextActionType: existing.nextActionType || inferNextActionType(row, stage, nextContact),
      nextContactAt: existing.nextContactAt ?? (nextContact ? nextContact.getTime() : null),
      assignedAdvisorEmail: existing.assignedAdvisorEmail || null,
      lastAdvisorEmail: existing.lastAdvisorEmail || tracking.advisorEmail || null,
      lastActionAt: existing.lastActionAt ?? (lastActionAt ? lastActionAt.getTime() : null),
      lastResultLabel: existing.lastResultLabel || tracking.resultLabel || String(getField(row, 'result') || '').trim() || null,
      attempts: Number(existing.attempts || tracking.attempts || getAttemptCount(row) || 0),
      isDuplicateCandidate: Boolean(existing.isDuplicateCandidate || false),
      duplicateGroupId: existing.duplicateGroupId || null,
      potentialMonthlyValue: Number.isFinite(Number(existing.potentialMonthlyValue)) ? Number(existing.potentialMonthlyValue) : null,
      probability: Number.isFinite(Number(existing.probability)) ? Number(existing.probability) : null
    };
    crm.stageLabel = CRM_STAGES[crm.stage]?.label || crm.stage;
    crm.temperatureLabel = CRM_TEMPERATURES[crm.temperature]?.label || crm.temperature;
    return crm;
  }

  function inferCrmStage(row) {
    const blob = [
      getField(row, 'status'),
      getField(row, 'result'),
      getField(row, ['Listado', 'Listado1']),
      getField(row, ['Comentario', 'Observaciones', 'Notas', 'Etapa', 'Gestion', 'GestiÃ³n'])
    ].map(normalizeText).join(' | ');
    const tracking = getTracking(row);
    const status = normalizeText(tracking.status || tracking.lastAction || tracking.resultLabel || '');

    if (/invalid|dato.*invalid|numero invalido|telefono invalido|datos invalidos/.test(`${blob} ${status}`)) return 'invalid';
    if (/matriculad|inscrit|\bactiv[oa]\b|estudiante activo|ya inicio|enrolled/.test(`${blob} ${status}`) && !/no matriculad|no inscrit|retirad/.test(blob)) return 'enrolled';
    if (/no interesad|no desea|no quiere|descartad|cerrad|perdid/.test(blob)) return 'not_interested';
    if (/excluido confirmado|archivad|bloquead|duplicad|retirad/.test(blob)) return 'archived';
    if (/pendiente.*pago|pago pendiente|payment/.test(blob)) return 'payment_pending';
    if (/clase.*realiz|trial done/.test(blob)) return 'trial_done';
    if (/clase.*agend|clase de prueba|trial|agendad/.test(blob)) return 'trial_scheduled';
    if (/interes|cotiz|precio|horario|pregunt/.test(blob)) return 'interested';
    if (/contactado|respondio|rescheduled|reprogramado/.test(`${blob} ${status}`)) return 'contacted';
    if (/reactiv|recover|recovered/.test(`${blob} ${status}`)) return 'reactivation';
    return 'new';
  }

  function calculateLeadScore(row, stage = inferCrmStage(row)) {
    const blob = [
      getField(row, 'channel'),
      getField(row, 'art'),
      getField(row, 'modality'),
      getField(row, 'plan'),
      getField(row, 'result'),
      getField(row, 'status'),
      getField(row, ['Comentario', 'Observaciones', 'Notas'])
    ].map(normalizeText).join(' | ');
    let score = 0;
    if (stage === 'trial_scheduled') score += 30;
    if (/referid/.test(blob)) score += 25;
    if (/horario|agenda|agend/.test(blob)) score += 25;
    if (/respondio|contactado|whatsapp/.test(blob) || getTracking(row).status === 'contacted') score += 20;
    if (String(getField(row, 'art') || '').trim()) score += 20;
    if (/precio|cotiz|valor|interes/.test(blob)) score += 15;
    if (String(getField(row, 'modality') || '').trim()) score += 15;
    if (normalizePhone(getField(row, 'phone')).length >= 7 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(getField(row, 'email') || '').trim())) score += 10;
    const attempts = getAttemptCount(row);
    if (/no respondio|no respuesta/.test(blob) && attempts >= 2) score -= 10;
    if (/no respondio|no respuesta/.test(blob) && attempts >= 3) score -= 20;
    if (stage === 'not_interested') score -= 30;
    if (stage === 'invalid') score -= 40;
    if (stage === 'enrolled') score = Math.max(score, 90);
    return Math.max(0, Math.min(100, score));
  }

  function scoreToTemperature(score) {
    const n = Number(score) || 0;
    if (n >= 80) return 'hot';
    if (n >= 50) return 'warm';
    if (n >= 20) return 'cold';
    return 'frozen';
  }

  function inferNextActionType(row, stage, nextContact) {
    if (['enrolled', 'archived', 'not_interested', 'invalid'].includes(stage)) return 'none';
    if (nextContact) return 'schedule';
    if (normalizePhone(getField(row, 'phone')).length >= 7) return 'whatsapp';
    return 'call';
  }

  function normalizePhone(value) {
    return String(value || '').replace(/\D/g, '');
  }

  function renderCrmStageBadge(stage) {
    const def = CRM_STAGES[stage] || CRM_STAGES.new;
    return `<span class="crm-badge ${escAttr(def.className)}">${esc(def.label)}</span>`;
  }

  function renderTemperatureBadge(temperature, score) {
    const def = CRM_TEMPERATURES[temperature] || CRM_TEMPERATURES.frozen;
    return `<span class="crm-badge ${escAttr(def.className)}">${esc(def.label)} <small>${Number(score || 0)}</small></span>`;
  }

  function renderCrmSummaryHTML(row) {
    const crm = normalizeCrm(row);
    const next = crm.nextContactAt ? formatYMD(toYMD(new Date(crm.nextContactAt))) : 'Sin fecha';
    const result = crm.lastResultLabel || 'Sin gestion registrada';
    return `<div class="crm-summary field">
      <div class="label">Estado comercial</div>
      <div class="value">
        ${renderCrmStageBadge(crm.stage)}
        ${renderTemperatureBadge(crm.temperature, crm.leadScore)}
        <span class="crm-meta">Proxima accion: ${esc(crm.nextActionType)} · ${esc(next)} · ${esc(result)}</span>
      </div>
    </div>`;
  }

  function findHeaderIncludes(row, keywords){
    return Object.keys(row || {}).find((key) => {
      const normalized = normalizeText(key);
      return keywords.every(word => normalized.includes(word));
    }) || '';
  }

  function normalizeText(value){
    return String(value ?? '')
      .normalize('NFD').replace(/\p{Diacritic}/gu, '')
      .toLowerCase()
      .replace(/[^\p{L}\p{N}]+/gu, ' ')
      .trim()
      .replace(/\s+/g, ' ');
  }

  function parseFlexibleDate(value){
    if (typeof value === 'number' && Number.isFinite(value)) return toYMD(new Date(value));
    if (value && typeof value.toDate === 'function') return toYMD(value.toDate());
    const s = String(value ?? '').trim();
    if (!s) return null;
    return parseYMD(s) || null;
  }

  function isExcludedFromFollowUp(row) {
    const fields = [
      'Estado', 'Status', 'Listado', 'Listado1', 'Matriculado', 'Observaciones', 'Notas',
      'Resultado', 'Gestion', 'GestiÃ³n', 'Etapa', 'Canal', 'Canal de comunicacion',
      'Canal de comunicaciÃ³n', 'Comentario'
    ];
    const blob = fields.map(name => normalizeText(getField(row, [name]))).filter(Boolean).join(' | ');
    if (!blob) return false;

    const excludedPatterns = [
      /matriculad[oa]/, /inscrit[oa]/, /\bactiv[oa]\b/, /estudiante activo/, /ya inicio/,
      /no interesad[oa]/, /no desea/, /no quiere/, /desistio/, /descartad[oa]/, /cerrad[oa]/, /perdid[oa]/,
      /numero invalido/, /telefono invalido/, /datos invalidos/, /numero no existe/, /no responde nunca/, /bloquead[oa]/,
      /duplicad[oa]/, /repetid[oa]/,
      /retirad[oa]/, /no disponible/, /fuera de cobertura/, /\berror\b/, /\bprueba\b/
    ];
    const negativeActive = /no matriculad[oa]|no inscrit[oa]|no activ[oa]/.test(blob);
    return excludedPatterns.some(rx => rx.test(blob)) && !negativeActive;
  }

  async function computeAndRenderAlerts() {
    if (!alertsEl && !todayViewEl) return;

    let cache = allRowsCache.get(currentSheet);
    const needsNetwork = !cache;
    if (needsNetwork) {
      startFauxProgress('Cargando "Hoy en Musicala"', 'Leyendo contactos de la base...', { start: 10, ceiling: 88 });
      status('Preparando "Hoy en Musicala"...');
    }
    try {
      if (!cache) {
        cache = await loadAllRowsForSheet(currentSheet);
        allRowsCache.set(currentSheet, cache);
      }

      const rows = cache.rows || [];
      // Detenemos la barra "animada": de aqui en adelante el avance es real por lotes.
      if (needsNetwork) {
        stopFauxProgressTimer();
        setProgress(60, `Organizando ${rows.length.toLocaleString()} contactos...`);
      }
      // Fase 1: tarjetas de "Hoy" (60% -> 80%)
      const onTodayProgress = needsNetwork
        ? (done, tot) => setProgress(60 + Math.round((done / Math.max(1, tot)) * 20), `Analizando contactos (${done.toLocaleString()} de ${tot.toLocaleString()})...`)
        : null;
      await renderTodayView(rows, onTodayProgress);
      // Fase 2: cola recomendada (80% -> 96%)
      const onQueueProgress = needsNetwork
        ? (done, tot) => setProgress(80 + Math.round((done / Math.max(1, tot)) * 16), `Priorizando seguimientos (${done.toLocaleString()} de ${tot.toLocaleString()})...`)
        : null;
      const result = await buildTopRecommendedQueue(rows, onQueueProgress);
      console.info('[top10 seguimiento]', result.debug);
      renderAlertsUI(result);
      if (needsNetwork) {
        finishFauxProgress('"Hoy en Musicala" listo.');
        status(`"Hoy en Musicala" listo: ${rows.length.toLocaleString()} registros analizados.`);
      }
    } catch (err) {
      if (needsNetwork) {
        stopFauxProgressTimer();
        hideProgress();
      }
      throw err;
    }
  }

  async function renderTodayView(rows, onProgress = null) {
    if (!todayViewEl) return;
    const groups = await buildTodayGroups(rows || [], onProgress);
    const defs = todayGroupDefs();
    const totalItems = defs.reduce((sum, def) => sum + groups[def.key].length, 0);

    if (!rows.length) {
      todayViewEl.classList.add('hidden');
      todayViewEl.innerHTML = '';
      return;
    }

    todayViewEl.classList.toggle('today-collapsed', !todayViewExpanded);
    todayViewEl.innerHTML = `
      <div class="today-head">
        <div>
          <span class="alert-eyebrow">Hoy</span>
          <h2>Hoy en Musicala</h2>
          <p>${todayViewExpanded ? 'Contactos accionables segun fechas, etapa CRM, prioridad y ultimo seguimiento.' : `${totalItems.toLocaleString()} contactos accionables detectados. Abre el panel para ver las tarjetas.`}</p>
        </div>
        <div class="today-head-actions">
          <div class="today-kpis">
            ${defs.slice(0, 4).map(def => `<span>${esc(def.label)} <strong>${groups[def.key].length.toLocaleString()}</strong></span>`).join('')}
          </div>
          <button id="btnToggleTodayView" class="btn today-toggle" type="button" aria-expanded="${todayViewExpanded ? 'true' : 'false'}">
            ${todayViewExpanded ? 'Minimizar' : 'Ver completo'}
          </button>
        </div>
      </div>
      ${todayViewExpanded ? `
        <div class="today-groups">
          ${defs.map(def => renderTodayGroup(def, groups[def.key])).join('')}
        </div>
        ${totalItems ? '' : '<div class="alert-empty">No hay contactos accionables para hoy con las reglas actuales.</div>'}
      ` : ''}
    `;
    todayViewEl.classList.remove('hidden');

    todayViewEl.querySelector('#btnToggleTodayView')?.addEventListener('click', () => {
      todayViewExpanded = !todayViewExpanded;
      renderTodayView(rows);
    });

    todayViewEl.querySelectorAll('[data-openid].today-open').forEach(btn => {
      btn.addEventListener('click', async () => {
        const row = await getRowById(btn.getAttribute('data-openid') || '');
        if (row) openModal(row);
        else alert('No se encontro la fila seleccionada.');
      });
    });
    todayViewEl.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', () => handleRecommendedAction(btn));
    });
  }

  function todayGroupDefs() {
    return [
      { key:'today', label:'Para contactar hoy' },
      { key:'overdue', label:'Vencidos' },
      { key:'new', label:'Nuevos sin gestionar' },
      { key:'highNoDate', label:'Alta prioridad sin fecha' },
      { key:'reactivation', label:'Reactivaciones' },
      { key:'trial', label:'Clases de prueba proximas' },
      { key:'payment', label:'Pendientes de pago' },
      { key:'invalid', label:'Datos invalidos para auditar' }
    ];
  }

  async function buildTodayGroups(rows, onProgress = null) {
    const today = toYMD(new Date());
    const limitPerGroup = 6;
    const groups = {
      today: [], overdue: [], new: [], highNoDate: [],
      reactivation: [], trial: [], payment: [], invalid: []
    };
    const add = (key, item) => {
      if (!groups[key] || groups[key].length >= limitPerGroup) return;
      groups[key].push(item);
    };

    await forEachChunked(rows, (row, index) => {
      if (isActiveOrEnrolled(row)) return;
      const crm = normalizeCrm(row);
      const next = getNextContactDate(row);
      const last = getLastManagedDate(row);
      const priority = normalizeText(getField(row, 'priority'));
      const rec = classifyRecommendedFollowUp(row, index, today);
      const item = { row, crm, rec, index };

      if (next && next.getTime() === today.getTime()) add('today', item);
      if (next && next < today) add('overdue', item);
      if (!last && crm.stage === 'new') add('new', item);
      if (!next && priority === 'alta') add('highNoDate', item);
      if (crm.stage === 'reactivation' || rec.bucket === 'reactivation_never_contacted' || rec.bucket === 'reactivation_old_client_or_warm') add('reactivation', item);
      if (crm.stage === 'trial_scheduled' || crm.stage === 'trial_done') add('trial', item);
      if (crm.stage === 'payment_pending') add('payment', item);
      if (crm.stage === 'invalid' || rec.bucket === 'excluded_audit') add('invalid', item);
    }, { onProgress });

    Object.keys(groups).forEach((key) => groups[key].sort(compareTodayItems));
    return groups;
  }

  function compareTodayItems(a, b) {
    const an = getNextContactDate(a.row);
    const bn = getNextContactDate(b.row);
    if (an && bn && an.getTime() !== bn.getTime()) return an - bn;
    if (an && !bn) return -1;
    if (!an && bn) return 1;
    return Number(b.crm.leadScore || 0) - Number(a.crm.leadScore || 0);
  }

  function renderTodayGroup(def, items) {
    return `
      <section class="today-group">
        <div class="today-group-head">
          <h3>${esc(def.label)}</h3>
          <span>${items.length.toLocaleString()}</span>
        </div>
        <div class="today-card-list">
          ${items.length ? items.map(renderTodayCard).join('') : '<div class="today-empty">Sin contactos</div>'}
        </div>
      </section>
    `;
  }

  function renderTodayCard(item) {
    const row = item.row;
    const crm = item.crm || normalizeCrm(row);
    const id = String(getField(row, 'id') || '').trim();
    const name = String(getField(row, 'name') || '').trim() || '(Sin nombre)';
    const phone = String(getField(row, 'phone') || '').trim();
    const phoneDigits = normalizePhone(phone);
    const waNum = phoneDigits.startsWith('57') ? phoneDigits.slice(2) : phoneDigits;
    const wa = phoneDigits.length >= 7 ? `https://wa.me/57${escAttr(waNum)}` : '';
    const art = getField(row, 'art') || 'Sin arte';
    const channel = getField(row, 'channel') || 'Sin canal';
    const advisor = getField(row, 'advisor') || 'Sin asesor';
    const next = crm.nextContactAt ? formatYMD(toYMD(new Date(crm.nextContactAt))) : 'Sin fecha';
    const result = crm.lastResultLabel || 'Sin gestion';

    return `
      <article class="today-card">
        <div class="today-card-main">
          <strong>${esc(name)}</strong>
          <span>${esc([phone, art, channel].filter(Boolean).join(' - '))}</span>
          <span>${esc(`Asesor: ${advisor} - Proxima: ${next} - ${result}`)}</span>
        </div>
        <div class="today-card-badges">
          ${renderCrmStageBadge(crm.stage)}
          ${renderTemperatureBadge(crm.temperature, crm.leadScore)}
        </div>
        <div class="today-actions">
          <button class="btn today-open" data-openid="${escAttr(id)}" type="button">Abrir</button>
          ${wa ? `<a class="btn btn-whatsapp-mini" href="${wa}" target="_blank" rel="noopener">WhatsApp</a>` : ''}
          <button class="btn alert-action" data-action="contacted" data-openid="${escAttr(id)}" type="button">Contactado</button>
          <button class="btn alert-action" data-action="no_response" data-openid="${escAttr(id)}" type="button">No respondio</button>
          <button class="btn alert-action" data-action="reschedule" data-openid="${escAttr(id)}" type="button">Reprogramar</button>
          <button class="btn alert-action" data-action="enrolled" data-openid="${escAttr(id)}" type="button">Matriculado</button>
        </div>
      </article>
    `;
  }

  async function buildTopRecommendedQueue(rows, onProgress = null) {
    const today = toYMD(new Date());
    const groups = createRecommendedGroups();
    const counts = createRecommendedCounts(rows.length);
    let skippedByCooldown = 0;

    await forEachChunked(rows, (row, index) => {
      const rec = classifyRecommendedFollowUp(row, index, today);
      counts[rec.bucket] = (counts[rec.bucket] || 0) + 1;
      if (rec.skippedByCooldown) {
        skippedByCooldown += 1;
        return;
      }
      if (groups[rec.bucket]) groups[rec.bucket].push({ row, rec, index });
    }, { onProgress });

    Object.values(groups).forEach((items) => items.sort(compareRecommendedItems));

    const selected = [];
    const used = new Set();
    const quotaTaken = {};
    const takeFrom = (bucket, amount) => {
      const source = groups[bucket] || [];
      let taken = 0;
      for (const item of source) {
        if (selected.length >= TOP_RECOMMENDED_LIMIT || taken >= amount) break;
        const key = stableRowKey(item.row, item.index);
        if (used.has(key)) continue;
        selected.push(item);
        used.add(key);
        quotaTaken[bucket] = (quotaTaken[bucket] || 0) + 1;
        taken += 1;
      }
    };

    takeFrom('new_recent', TOP_RECOMMENDED_QUOTAS.new_recent);
    takeFrom('due_today', 1);
    takeFrom('overdue_recent', Math.max(0, TOP_RECOMMENDED_QUOTAS.due_recent - (quotaTaken.due_today || 0)));
    takeFrom('high_no_date', TOP_RECOMMENDED_QUOTAS.high_no_date);
    takeFrom('reactivation_never_contacted', TOP_RECOMMENDED_QUOTAS.reactivation_never_contacted);
    takeFrom('reactivation_old_client_or_warm', TOP_RECOMMENDED_QUOTAS.reactivation_old_client_or_warm);
    takeFrom('excluded_audit', TOP_RECOMMENDED_QUOTAS.excluded_audit);
    takeFrom('general_rotation', TOP_RECOMMENDED_QUOTAS.general_rotation);

    const refillOrder = [
      'due_today', 'overdue_recent', 'new_recent', 'reactivation_never_contacted',
      'reactivation_old_client_or_warm', 'high_no_date', 'excluded_audit',
      'general_rotation', 'upcoming'
    ];
    while (selected.length < TOP_RECOMMENDED_LIMIT) {
      const before = selected.length;
      refillOrder.forEach(bucket => takeFrom(bucket, 1));
      if (selected.length === before) break;
    }

    const coverage = buildCoverageStats(rows, today);
    return {
      top: selected.slice(0, TOP_RECOMMENDED_LIMIT),
      groups,
      counts: { ...counts, skippedByCooldown },
      coverage,
      debug: {
        ...counts,
        skippedByCooldown,
        top10: selected.length,
        quotaTaken,
        coverage
      }
    };
  }

  function renderAlertsUI(result){
    if (!alertsEl) return;
    const queue = result.top || [];
    const current = queue[0] || null;
    const counts = result.counts || {};
    const coverage = result.coverage || {};

    const makeItem = (item) => {
      const { row, rec } = item;
      const name = String(getField(row, ['Nombre']) || '').trim() || '(Sin nombre)';
      const phone = String(getField(row, ['Celular/TelÃ©fono', 'Celular/Telefono', 'Telefono', 'TelÃ©fono']) || '').trim();
      const canal = getField(row, ['Canal de comunicacion', 'Canal de comunicaciÃ³n', 'Canal']) || 'Sin canal';
      const asesor = getField(row, ['Asesor', 'Responsable']);
      const id = String(getField(row, ['ID']) || '').trim();
      const dateTxt = recommendedDateText(rec);
      const auditHint = rec.bucket === 'excluded_audit' ? '<span class="meta warning">Revision de excluido: verificar motivo antes de contactar.</span>' : '';

      return `
        <div class="alert-item alert-recommended alert-${escAttr(rec.bucket)}">
          <div class="alert-person">
            <strong>${esc(name)}</strong>
            <span class="meta">${esc([canal, phone, asesor ? `Asesor: ${asesor}` : ''].filter(Boolean).join(' - '))}</span>
            <span class="meta">${esc(rec.reason)}</span>
            ${dateTxt ? `<span class="meta">${esc(dateTxt)}</span>` : ''}
            ${auditHint}
          </div>
          <span class="mini-priority">${esc(rec.label)}</span>
          <div class="alert-actions">
            <button class="btn alert-open" data-openid="${escAttr(id)}">Abrir</button>
            <button class="btn alert-action" data-action="contacted" data-openid="${escAttr(id)}">Contactado</button>
            <button class="btn alert-action" data-action="no_response" data-openid="${escAttr(id)}">No respondio</button>
            <button class="btn alert-action" data-action="reschedule" data-openid="${escAttr(id)}">Reprogramar</button>
            <button class="btn alert-action" data-action="invalid_data" data-openid="${escAttr(id)}">Datos invalidos</button>
            <button class="btn alert-action" data-action="confirm_excluded" data-openid="${escAttr(id)}">Excluir confirmado</button>
            <button class="btn alert-action" data-action="enrolled" data-openid="${escAttr(id)}">Ya matriculado</button>
            <button class="btn alert-action" data-action="recover" data-openid="${escAttr(id)}">Recuperar</button>
            <button class="btn alert-action" data-action="skip" data-openid="${escAttr(id)}">Saltar</button>
          </div>
        </div>
      `;
    };

    alertsEl.innerHTML = `
      <div class="alert-card">
        <div class="alert-title">
          <div>
            <span class="alert-eyebrow">Seguimiento guiado</span>
            <h2>Siguiente contacto recomendado</h2>
            <p class="alert-subtitle">El sistema rota contactos nuevos, vencidos, antiguos y casos por auditar para avanzar sobre toda la base.</p>
          </div>
          <div class="pills">
            <span class="pill urgent">Contacto recomendado listo</span>
            <span class="pill soon">Cola interna activa: ${queue.length.toLocaleString()}</span>
            <span class="pill soon">${Number(counts.new_recent || 0).toLocaleString()} nuevos disponibles</span>
            <span class="pill review">${Number((counts.reactivation_never_contacted || 0) + (counts.reactivation_old_client_or_warm || 0)).toLocaleString()} reactivacion</span>
          </div>
        </div>
        <div class="alert-summary">
          ${summaryPill('Vencidos recientes', counts.overdue_recent)}
          ${summaryPill('Hoy', counts.due_today)}
          ${summaryPill('Alta sin fecha', counts.high_no_date)}
          ${summaryPill('Excluidos por auditar', counts.excluded_audit)}
          ${summaryPill('Rotacion general', counts.general_rotation)}
          ${summaryPill('Pendientes elegibles', coverage.eligible)}
          ${summaryPill('Revisados 7 dias', coverage.reviewed7)}
          ${summaryPill('Revisados 30 dias', coverage.reviewed30)}
        </div>
        ${current ? `
          <div class="alert-list alert-list-current">
            ${makeItem(current)}
          </div>
        ` : `
          <div class="alert-empty">No hay un contacto recomendado disponible en este momento.</div>
        `}
      </div>
    `;
    alertsEl.classList.remove('hidden');

    alertsEl.querySelectorAll('[data-openid].alert-open').forEach(btn => {
      btn.addEventListener('click', async () => {
        const row = await getRowById(btn.getAttribute('data-openid') || '');
        if (row) {
          openModal(row);
          setTimeout(()=>document.getElementById('btnEdit')?.click(), 0);
        } else {
          alert('No se encontro la fila seleccionada.');
        }
      });
    });

    alertsEl.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', () => handleRecommendedAction(btn));
    });
  }

  function createRecommendedGroups(){
    return {
      new_recent: [], due_today: [], overdue_recent: [], high_no_date: [],
      reactivation_never_contacted: [], reactivation_old_client_or_warm: [],
      excluded_audit: [], general_rotation: [], upcoming: [], scheduled_future: [], ok: [], excluded: []
    };
  }

  function createRecommendedCounts(totalBase){
    return {
      totalBase, top10: 0, new_recent: 0, due_today: 0, overdue_recent: 0,
      high_no_date: 0, reactivation_never_contacted: 0,
      reactivation_old_client_or_warm: 0, excluded_audit: 0, general_rotation: 0,
      upcoming: 0, scheduled_future: 0, ok: 0, excluded: 0, skippedByCooldown: 0
    };
  }

  function classifyRecommendedFollowUp(row, index, today){
    const tracking = getTracking(row);
    const trackedStatus = normalizeText(tracking.status);
    const trackedLastManaged = parseFlexibleDate(tracking.lastManagedAt || tracking.updatedAt);
    const trackedNextContact = parseFlexibleDate(tracking.nextContactAt);
    const trackedAttempts = Number(tracking.attempts || 0);
    const trackedBase = {
      rowIndex:index,
      attempts:Number.isFinite(trackedAttempts) ? trackedAttempts : 0,
      rotation:stableDailyRotation(row, index, today),
      lastManaged:trackedLastManaged,
      relevantDate:null,
      relevantDateLabel:'',
      skippedByCooldown:false
    };

    if (trackedStatus) {
      if (['enrolled', 'invalid data', 'invalid_data', 'excluded'].includes(trackedStatus)) {
        return { ...trackedBase, bucket:'excluded', label:tracking.resultLabel || 'Fuera de cola', reason:'Seguimiento Firebase: fuera de la cola normal' };
      }
      if (trackedStatus === 'contacted') {
        const days = trackedLastManaged ? daysBetween(trackedLastManaged, today) : 0;
        if (!trackedLastManaged || days < CONTACTED_COOLDOWN_DAYS) {
          return { ...trackedBase, bucket:'ok', label:'Contactado', reason:'Contactado recientemente en Firebase', skippedByCooldown:true };
        }
      }
      if (trackedStatus === 'rescheduled') {
        if (!trackedNextContact) return { ...trackedBase, bucket:'scheduled_future', label:'Reprogramado', reason:'Reprogramado sin fecha valida en Firebase' };
        const daysDue = daysBetween(trackedNextContact, today);
        if (trackedNextContact.getTime() === today.getTime()) return { ...trackedBase, bucket:'due_today', label:'Hoy', reason:'Reprogramado para hoy en Firebase', relevantDate:trackedNextContact, relevantDateLabel:'Para contactar' };
        if (trackedNextContact < today && daysDue <= 30) return { ...trackedBase, bucket:'overdue_recent', label:'Vencido reciente', reason:'Reprogramacion vencida en Firebase', relevantDate:trackedNextContact, relevantDateLabel:'Para contactar' };
        if (trackedNextContact < today) return { ...trackedBase, bucket:'reactivation_old_client_or_warm', label:'Reactivacion', reason:'Reprogramacion antigua vencida en Firebase', relevantDate:trackedNextContact, relevantDateLabel:'Para contactar' };
        if (trackedNextContact <= addDays(today, 7)) return { ...trackedBase, bucket:'upcoming', label:'Proximo', reason:'Reprogramado en los proximos 7 dias', relevantDate:trackedNextContact, relevantDateLabel:'Para contactar' };
        return { ...trackedBase, bucket:'scheduled_future', label:'Programado', reason:'Reprogramado a futuro en Firebase', relevantDate:trackedNextContact, relevantDateLabel:'Para contactar' };
      }
      if (trackedStatus === 'no response' || trackedStatus === 'no_response') {
        const days = trackedLastManaged ? daysBetween(trackedLastManaged, today) : 0;
        if (days < NO_RESPONSE_COOLDOWN_DAYS) return { ...trackedBase, bucket:'ok', label:'No respondio', reason:'No respondio recientemente en Firebase', skippedByCooldown:true };
      }
      if (trackedStatus === 'skipped') {
        const days = trackedLastManaged ? daysBetween(trackedLastManaged, today) : 0;
        if (days < SKIPPED_COOLDOWN_DAYS) return { ...trackedBase, bucket:'ok', label:'Saltado', reason:'Saltado recientemente en Firebase', skippedByCooldown:true };
      }
    }

    if (isActiveOrEnrolled(row)) {
      return { rowIndex:index, attempts:getAttemptCount(row), rotation:stableDailyRotation(row, index, today), lastManaged:getLastManagedDate(row), relevantDate:null, relevantDateLabel:'', skippedByCooldown:false, bucket:'ok', label:'Matriculado', reason:'Matriculado o activo: fuera de seguimiento' };
    }

    const excluded = isExcludedFromFollowUp(row);
    const auditDue = isExcludedAuditDue(row, today);
    const nextContact = getNextContactDate(row);
    const lastContact = getLastContactDate(row);
    const created = getCreatedDate(row);
    const lastManaged = getLastManagedDate(row);
    const attempts = getAttemptCount(row);
    const priority = normalizeText(getField(row, ['Prioridad']));
    const cooldown = getCooldownInfo(row, today);
    const rotation = stableDailyRotation(row, index, today);

    const base = { rowIndex:index, attempts, rotation, lastManaged, relevantDate:null, relevantDateLabel:'', skippedByCooldown:false };

    if (excluded) {
      if (!auditDue) return { ...base, bucket:'excluded', label:'Excluido', reason:'Excluido confirmado o en cooldown de auditoria' };
      if (cooldown.active) return { ...base, bucket:'excluded_audit', label:'Auditoria excluido', reason:`Revisar exclusion despues de cooldown (${cooldown.reason})`, skippedByCooldown:true };
      return { ...base, bucket:'excluded_audit', label:'Revision de excluido', reason:'Auditar si la exclusion sigue vigente', relevantDate:getNextAuditDate(row), relevantDateLabel:'Proxima auditoria' };
    }

    if (cooldown.active) return { ...base, bucket:'ok', label:'En cooldown', reason:cooldown.reason, skippedByCooldown:true };

    if (nextContact) {
      const daysDue = daysBetween(nextContact, today);
      if (nextContact.getTime() === today.getTime()) return { ...base, bucket:'due_today', label:'Hoy', reason:'Seguimiento programado para hoy', relevantDate:nextContact, relevantDateLabel:'Para contactar' };
      if (nextContact < today && daysDue <= 30) return { ...base, bucket:'overdue_recent', label:'Vencido reciente', reason:'Fecha de seguimiento vencida hace maximo 30 dias', relevantDate:nextContact, relevantDateLabel:'Para contactar' };
      if (nextContact < today && daysDue > 30) return { ...base, bucket:'reactivation_old_client_or_warm', label:'Reactivacion', reason:'Fecha vencida antigua: reactivar sin tratar como urgente', relevantDate:nextContact, relevantDateLabel:'Para contactar' };
      if (nextContact <= addDays(today, 7)) return { ...base, bucket:'upcoming', label:'Proximo', reason:'Seguimiento programado en los proximos 7 dias', relevantDate:nextContact, relevantDateLabel:'Para contactar' };
      return { ...base, bucket:'scheduled_future', label:'Programado', reason:'Fecha futura mayor a 7 dias', relevantDate:nextContact, relevantDateLabel:'Para contactar' };
    }

    if (priority === 'alta') return { ...base, bucket:'high_no_date', label:'Alta sin fecha', reason:'Prioridad alta sin fecha de seguimiento' };

    if (!lastContact) {
      if (created && daysBetween(created, today) <= 30) return { ...base, bucket:'new_recent', label:'Nuevo sin contactar', reason:'Contacto reciente sin primer contacto', relevantDate:created, relevantDateLabel:'Creado' };
      return { ...base, bucket:'reactivation_never_contacted', label:'Reactivacion sin contacto', reason:created ? 'Contacto antiguo nunca contactado' : 'Sin contacto registrado ni fecha de creacion', relevantDate:created, relevantDateLabel:created ? 'Creado' : '' };
    }

    const daysLast = daysBetween(lastContact, today);
    if (daysLast > 30 || hasWarmSignal(row)) return { ...base, bucket:'reactivation_old_client_or_warm', label:'Reactivacion', reason:'Contacto antiguo o con interes previo para reciclar', relevantDate:lastContact, relevantDateLabel:'Ultimo contacto' };
    if (daysLast >= 7) return { ...base, bucket:'general_rotation', label:'Rotacion general', reason:'Contacto elegible para cobertura progresiva', relevantDate:lastContact, relevantDateLabel:'Ultimo contacto' };
    return { ...base, bucket:'ok', label:'Al dia', reason:'Gestionado recientemente', relevantDate:lastContact, relevantDateLabel:'Ultimo contacto' };
  }

  function compareRecommendedItems(a, b){
    const am = a.rec.lastManaged ? a.rec.lastManaged.getTime() : 0;
    const bm = b.rec.lastManaged ? b.rec.lastManaged.getTime() : 0;
    if (am !== bm) return am - bm;
    if (a.rec.attempts !== b.rec.attempts) return a.rec.attempts - b.rec.attempts;
    const ad = a.rec.relevantDate;
    const bd = b.rec.relevantDate;
    if (ad && bd && ad.getTime() !== bd.getTime()) {
      if (a.rec.bucket === 'new_recent') return bd - ad;
      return ad - bd;
    }
    if (ad && !bd) return -1;
    if (!ad && bd) return 1;
    return a.rec.rotation - b.rec.rotation;
  }

  function stableDailyRotation(row, index, date){
    const day = formatYMD(date);
    const raw = `${day}|${stableRowKey(row, index)}`;
    let hash = 2166136261;
    for (let i = 0; i < raw.length; i += 1) {
      hash ^= raw.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0) / 4294967295;
  }

  function stableRowKey(row, index){
    return [
      getField(row, ['ID']),
      getField(row, ['Nombre']),
      getField(row, ['Celular/TelÃ©fono', 'Celular/Telefono', 'Telefono', 'TelÃ©fono']),
      getField(row, ['Correo Electronico', 'Correo ElectrÃ³nico', 'Email']),
      index
    ].map(v => normalizeText(v)).join('|');
  }

  function recommendedDateText(rec){
    if (!rec?.relevantDate || !rec?.relevantDateLabel) return '';
    return `${rec.relevantDateLabel}: ${formatYMD(rec.relevantDate)}`;
  }

  function getLastManagedDate(row){
    const tracking = getTracking(row);
    const trackedDate = parseFlexibleDate(tracking.lastManagedAt || tracking.updatedAt);
    if (trackedDate) return trackedDate;
    return firstDateFrom(row, ['Fecha ultima gestion', 'Fecha Ãºltima gestiÃ³n', 'Fecha ultima revision', 'Fecha Ãºltima revisiÃ³n', 'Ultima gestion', 'Ãšltima gestiÃ³n']);
  }

  function getNextAuditDate(row){
    return firstDateFrom(row, ['Fecha proxima auditoria', 'Fecha prÃ³xima auditorÃ­a', 'Proxima auditoria', 'PrÃ³xima auditorÃ­a']);
  }

  function getAttemptCount(row){
    const tracking = getTracking(row);
    const trackedAttempts = Number(tracking.attempts);
    if (Number.isFinite(trackedAttempts) && trackedAttempts > 0) return trackedAttempts;
    const raw = getField(row, ['Intentos', 'Numero de intentos', 'NÃºmero de intentos']);
    const n = Number(String(raw || '').replace(/[^\d]/g, ''));
    return Number.isFinite(n) ? n : 0;
  }

  function getCooldownInfo(row, today){
    const result = normalizeText(getField(row, ['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado']));
    const state = normalizeText(getField(row, ['Estado de seguimiento', 'Estado', 'Status']));
    const last = getLastManagedDate(row);
    if (!last) return { active:false, reason:'' };
    const days = daysBetween(last, today);
    if (/no respondio|no respuesta/.test(result) && days < NO_RESPONSE_COOLDOWN_DAYS) return { active:true, reason:'No respondio recientemente' };
    if (/contactado/.test(result) && days < CONTACTED_COOLDOWN_DAYS) return { active:true, reason:'Contactado recientemente' };
    if (/saltado/.test(result) && days < SKIPPED_COOLDOWN_DAYS) return { active:true, reason:'Saltado recientemente' };
    if (/datos invalidos/.test(result) && days < EXCLUDED_AUDIT_DAYS) return { active:true, reason:'Datos invalidos en auditoria' };
    if (/excluido confirmado/.test(state) && days < CONFIRMED_EXCLUDED_AUDIT_DAYS) return { active:true, reason:'Exclusion confirmada reciente' };
    return { active:false, reason:'' };
  }

  function isExcludedAuditDue(row, today){
    const nextAudit = getNextAuditDate(row);
    if (nextAudit) return nextAudit <= today;
    const state = normalizeText(getField(row, ['Estado de seguimiento', 'Estado', 'Status']));
    const last = getLastManagedDate(row);
    if (/excluido confirmado/.test(state) && last) return daysBetween(last, today) >= CONFIRMED_EXCLUDED_AUDIT_DAYS;
    return true;
  }

  function isActiveOrEnrolled(row){
    const fields = ['Estado', 'Status', 'Listado', 'Listado1', 'Matriculado', 'Resultado', 'Gestion', 'GestiÃ³n', 'Etapa'];
    const blob = fields.map(name => normalizeText(getField(row, [name]))).filter(Boolean).join(' | ');
    if (/no matriculad[oa]|no inscrit[oa]|retirad[oa]/.test(blob)) return false;
    return /matriculad[oa]|inscrit[oa]|\bactiv[oa]\b|estudiante activo|ya inicio/.test(blob);
  }

  function hasWarmSignal(row){
    const blob = ['Comentario', 'Observaciones', 'Notas', 'Resultado', 'Gestion', 'GestiÃ³n', 'Etapa', 'Clase', 'Curso/Plan']
      .map(name => normalizeText(getField(row, [name]))).join(' | ');
    return /clase|prueba|interes|interesad|cotiz|pregunto|agenda|agendo|volvio|retomar/.test(blob);
  }

  function buildCoverageStats(rows, today){
    const stats = { reviewed7:0, reviewed30:0, unreviewed:0, eligible:0, confirmedExcluded:0 };
    rows.forEach((row, index) => {
      const state = normalizeText(getField(row, ['Estado de seguimiento', 'Estado', 'Status']));
      if (/excluido confirmado/.test(state)) stats.confirmedExcluded += 1;
      const rec = classifyRecommendedFollowUp(row, index, today);
      if (!['excluded', 'scheduled_future', 'ok'].includes(rec.bucket)) stats.eligible += 1;
      const last = getLastManagedDate(row);
      if (!last) stats.unreviewed += 1;
      else {
        const days = daysBetween(last, today);
        if (days <= 7) stats.reviewed7 += 1;
        if (days <= 30) stats.reviewed30 += 1;
      }
    });
    return stats;
  }

  async function handleRecommendedAction(btn){
    const id = btn.getAttribute('data-openid') || '';
    const action = btn.getAttribute('data-action') || '';
    const row = await getRowById(id);
    if (!row) {
      alert('No se encontro la fila seleccionada.');
      return;
    }

    const extra = {};
    if (action === 'reschedule') {
      const value = prompt('Fecha para contactar (AAAA-MM-DD):', formatYMD(addDays(toYMD(new Date()), 3)));
      if (!value) return;
      extra.followUpDate = value;
    }
    if (action === 'confirm_excluded') {
      const reason = prompt('Motivo de exclusion:', 'No requiere seguimiento');
      if (reason === null) return;
      extra.reason = reason;
    }
    if (action === 'skip') {
      const reason = prompt('Motivo para saltar este contacto:', 'No es buen momento');
      if (reason === null) return;
      extra.reason = reason;
    }

    const actionScope = btn.closest('#todayView, #alerts') || document;
    const actionButtons = Array.from(actionScope.querySelectorAll('.alert-action') || []);
    const originalText = btn.textContent;
    actionButtons.forEach(actionBtn => { actionBtn.disabled = true; });
    btn.textContent = 'Guardando...';
    status('Guardando seguimiento...');
    try {
      const updated = await saveTrackingToFirebase(currentSheet, row, action, extra);
      upsertRowInLocalState(updated);
      renderCurrentPage();
      scheduleFollowUpRefresh(150);
      status('Seguimiento guardado.');
    } catch (err) {
      console.error(err);
      alert(`No se pudo guardar la gestion: ${err?.message || err}`);
      status(`Error al guardar seguimiento: ${err?.message || err}`);
      actionButtons.forEach(actionBtn => { actionBtn.disabled = false; });
      btn.textContent = originalText;
    }
  }

  function buildManagedRow(row, action, extra = {}){
    const out = { ...row };
    const now = currentDateTimeLocal();
    const today = toYMD(new Date());
    const set = (fieldOrNames, value) => setField(out, fieldOrNames, value);
    const attempts = getAttemptCount(row) + 1;

    set(['Fecha ultima gestion', 'Fecha Ãºltima gestiÃ³n', 'Ultima gestion', 'Ãšltima gestiÃ³n'], now);
    set(['Asesor ultima gestion', 'Asesor Ãºltima gestiÃ³n'], authUser?.email || '');
    set(['Intentos', 'Numero de intentos', 'NÃºmero de intentos'], String(attempts));

    if (action === 'contacted') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'Contactado');
      set(['Fecha y hora de contacto'], now);
    } else if (action === 'no_response') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'No respondio');
      set(['Fecha para contactar'], formatYMD(addDays(today, NO_RESPONSE_COOLDOWN_DAYS)));
    } else if (action === 'reschedule') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'Reprogramado');
      set(['Fecha para contactar'], extra.followUpDate || formatYMD(addDays(today, NO_RESPONSE_COOLDOWN_DAYS)));
    } else if (action === 'invalid_data') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'Datos invalidos');
      set(['Estado de seguimiento', 'Estado', 'Status'], 'Datos invalidos');
      set(['Fecha proxima auditoria', 'Fecha prÃ³xima auditorÃ­a'], formatYMD(addDays(today, EXCLUDED_AUDIT_DAYS)));
    } else if (action === 'confirm_excluded') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'Exclusion confirmada');
      set(['Estado de seguimiento', 'Estado', 'Status'], 'Excluido confirmado');
      set(['Motivo de exclusion', 'Motivo de exclusiÃ³n'], extra.reason || 'No requiere seguimiento');
      set(['Fecha proxima auditoria', 'Fecha prÃ³xima auditorÃ­a'], formatYMD(addDays(today, CONFIRMED_EXCLUDED_AUDIT_DAYS)));
    } else if (action === 'enrolled') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'Ya matriculado');
      set(['Estado de seguimiento', 'Estado', 'Status', 'Matriculado'], 'Matriculado / Activo');
    } else if (action === 'recover') {
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], 'Recuperado para seguimiento');
      set(['Estado de seguimiento', 'Estado', 'Status'], 'Recuperado');
      set(['Fecha para contactar'], formatYMD(today));
    } else if (action === 'skip') {
      set(['Fecha ultima revision', 'Fecha Ãºltima revisiÃ³n'], now);
      set(['Resultado gestion', 'Resultado gestiÃ³n', 'Resultado'], `Saltado: ${extra.reason || 'Sin motivo'}`);
    }

    return attachFirebaseMeta(out, {
      data: out,
      tracking: buildTrackingUpdate(row, action, extra),
      crm: normalizeCrm(out),
      source: row.__source || null
    }, row.__rowDocId);
  }

  function buildTrackingUpdate(row, action, extra = {}){
    const previous = cloneTracking(getTracking(row));
    const now = Date.now();
    const today = toYMD(new Date());
    const attempts = Number(previous.attempts || getAttemptCount(row) || 0) + 1;
    const base = {
      ...previous,
      lastManagedAt: now,
      advisorEmail: authUser?.email || '',
      attempts,
      updatedAt: now
    };

    if (action === 'contacted') {
      return { ...base, status:'contacted', resultLabel:'Contactado', lastAction:'contacted', lastContactAt:now, nextContactAt:null, excludedReason:null };
    }
    if (action === 'no_response') {
      return { ...base, status:'no_response', resultLabel:'No respondio', lastAction:'no_response', nextContactAt:addDays(today, NO_RESPONSE_COOLDOWN_DAYS).getTime(), excludedReason:null };
    }
    if (action === 'reschedule') {
      const followUp = parseFlexibleDate(extra.followUpDate) || addDays(today, NO_RESPONSE_COOLDOWN_DAYS);
      return { ...base, status:'rescheduled', resultLabel:'Reprogramado', lastAction:'rescheduled', nextContactAt:followUp.getTime(), excludedReason:null };
    }
    if (action === 'invalid_data') {
      return { ...base, status:'invalid_data', resultLabel:'Datos invalidos', lastAction:'invalid_data', nextContactAt:null, auditAt:addDays(today, EXCLUDED_AUDIT_DAYS).getTime(), excludedReason:'Datos invalidos' };
    }
    if (action === 'confirm_excluded') {
      return { ...base, status:'excluded', resultLabel:'Exclusion confirmada', lastAction:'excluded', nextContactAt:null, auditAt:addDays(today, CONFIRMED_EXCLUDED_AUDIT_DAYS).getTime(), excludedReason:extra.reason || 'No requiere seguimiento' };
    }
    if (action === 'enrolled') {
      return { ...base, status:'enrolled', resultLabel:'Ya matriculado', lastAction:'enrolled', lastContactAt:now, nextContactAt:null, excludedReason:null };
    }
    if (action === 'recover') {
      return { ...base, status:'recovered', resultLabel:'Recuperado para seguimiento', lastAction:'recovered', nextContactAt:today.getTime(), excludedReason:null };
    }
    if (action === 'skip') {
      return { ...base, status:'skipped', resultLabel:'Saltado', lastAction:'skipped', nextContactAt:addDays(today, SKIPPED_COOLDOWN_DAYS).getTime(), excludedReason:extra.reason || null };
    }
    return { ...base, status:'pending', resultLabel:'Pendiente', lastAction:action || 'pending' };
  }

  function buildContactSnapshot(row) {
    return {
      nombre: getField(row, 'name') || null,
      telefono: getField(row, 'phone') || null,
      correo: getField(row, 'email') || null,
      canal: getField(row, 'channel') || null,
      asesor: getField(row, 'advisor') || null,
      arte: getField(row, 'art') || null
    };
  }

  function renderCurrentPage() {
    renderTable(currentHeaders, currentRows);
  }

  function upsertRowInLocalState(row) {
    const idKey = keyColumnInUse || 'ID';
    const id = String(row?.[idKey] ?? row?.ID ?? '').trim();
    if (!id) return;
    const same = (candidate) => String(candidate?.[idKey] ?? candidate?.ID ?? '').trim() === id;
    const replaceIn = (rows) => {
      const list = Array.isArray(rows) ? rows : [];
      const idx = list.findIndex(same);
      if (idx >= 0) list[idx] = row;
      else list.push(row);
      return list;
    };

    currentRows = replaceIn(currentRows);
    displayedRows = replaceIn(displayedRows);

    const cache = allRowsCache.get(currentSheet) || { headers: currentHeaders, rows: [], total: 0, ts: Date.now(), source: 'firebase' };
    cache.rows = replaceIn(cache.rows);
    cache.total = Math.max(Number(cache.total || 0), cache.rows.length);
    cache.ts = Date.now();
    allRowsCache.set(currentSheet, cache);
  }

  function setExistingField(row, names, value){
    const key = findExistingKey(row, names);
    if (key) row[key] = value;
  }

  function findExistingKey(row, names){
    const direct = names.find(name => Object.prototype.hasOwnProperty.call(row, name));
    if (direct) return direct;
    const wanted = names.map(normalizeText);
    return Object.keys(row || {}).find(k => wanted.includes(normalizeText(k))) || '';
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
