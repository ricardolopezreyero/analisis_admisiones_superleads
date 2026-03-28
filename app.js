const state = {
  raw_rows: [],
  rows: [],
  filtered_rows: [],
  headers: [],
  col: {},
  institution_name: '',
  table: {
    page: 1,
    page_size: 10,
    is_collapsed: true
  },
  filters: {
    assigned: 'Todos',
    phase: 'Todas',
    level: 'Todos',
    time: 'Todo',
    date_from: '',
    date_to: ''
  },
  ui: {
    is_processing: false,
    analysis_timer_ids: []
  }
};

const drop_zone = document.getElementById('drop-zone');
const csv_input = document.getElementById('csv-file');
const btn_pick_csv = document.getElementById('btn-pick-csv');
const dashboard_title = document.getElementById('dashboard-title');
const hero_subtitle = document.getElementById('hero-subtitle');
const institution_hero_box = document.getElementById('institution-hero-box');
const institution_input = document.getElementById('institution-name');
const btn_main_action = document.getElementById('btn-main-action');
const controls = document.getElementById('controls');
const charts_wrap = document.getElementById('charts');
const table_wrap = document.getElementById('table-wrap');
const records_wrap = document.getElementById('records-wrap');
const records_grid = document.getElementById('records-grid');
const kpi_grid = document.getElementById('kpi-grid');
const insights_box = document.getElementById('insights');
const analysis_sequence = document.getElementById('analysis-sequence');
const analysis_title = document.getElementById('analysis-title');
const analysis_message = document.getElementById('analysis-message');
const analysis_progress_bar = document.getElementById('analysis-progress-bar');
const analysis_steps = [
  document.getElementById('analysis-step-1'),
  document.getElementById('analysis-step-2'),
  document.getElementById('analysis-step-3'),
  document.getElementById('analysis-step-4'),
  document.getElementById('analysis-step-5')
];

const safe_storage = {
  get(key) {
    try {
      return window.localStorage.getItem(key);
    } catch (_error) {
      return '';
    }
  },
  set(key, value) {
    try {
      window.localStorage.setItem(key, value);
    } catch (_error) {
      // En algunos contenedores el storage está bloqueado.
    }
  }
};

const filter_assigned = document.getElementById('filter-assigned');
const filter_phase = document.getElementById('filter-phase');
const filter_level = document.getElementById('filter-level');
const filter_time = document.getElementById('filter-time');
const time_custom_range = document.getElementById('time-custom-range');
const filter_date_from = document.getElementById('filter-date-from');
const filter_date_to = document.getElementById('filter-date-to');
const btn_reset = document.getElementById('btn-reset');
const btn_export_csv = document.getElementById('btn-export-csv');
const btn_export_report = document.getElementById('btn-export-report');
const btn_export_pdf = document.getElementById('btn-export-pdf');
const btn_toggle_table = document.getElementById('btn-toggle-table');
const table_scroll = document.getElementById('table-scroll');
const table_pagination = document.getElementById('table-pagination');
const btn_page_prev = document.getElementById('btn-page-prev');
const btn_page_next = document.getElementById('btn-page-next');
const table_page_info = document.getElementById('table-page-info');

const phase_order = [
  'Nueva Oportunidad',
  'En Proceso de Contactación',
  'Contactado',
  'Citas',
  'Aceptado',
  'Inscrito',
  'Ganado',
  'Perdido'
];

const chart_colors = {
  primary: '#0039C8',
  secondary: '#2A89FB',
  accent: '#1DB2FC',
  success: '#56EF9F',
  success_hover: '#2BC878',
  neutral: '#EEF2FF',
  muted: '#9FB7E8'
};

const phase_palette = {
  'Nueva Oportunidad': '#1DB2FC',
  'En Proceso de Contactación': '#2A89FB',
  Contactado: '#0039C8',
  Citas: '#56EF9F',
  Aceptado: '#56EF9F',
  Inscrito: '#2BC878',
  Ganado: '#56EF9F',
  Perdido: '#EEF2FF',
  'Sin fase': '#EEF2FF'
};

init();

function init() {
  const remembered_name = safe_storage.get('dashboard_institution_name') || '';
  if (remembered_name) {
    institution_input.value = remembered_name;
  }
  refresh_main_action_button();
  sync_drop_zone_state();

  drop_zone.addEventListener('click', () => {
    if (state.ui.is_processing) {
      return;
    }
    if (!state.institution_name) {
      alert('Primero guarda el nombre de la institución.');
      institution_input.focus();
      return;
    }
    open_file_picker();
  });

  drop_zone.addEventListener('keydown', (event) => {
    if (state.ui.is_processing) {
      return;
    }
    if (event.key !== 'Enter' && event.key !== ' ') {
      return;
    }
    event.preventDefault();
    if (!state.institution_name) {
      alert('Primero guarda el nombre de la institución.');
      institution_input.focus();
      return;
    }
    open_file_picker();
  });
  csv_input.addEventListener('change', (event) => {
    const file = event.target.files?.[0];
    if (file) {
      parse_csv(file);
    }
  });

  if (btn_pick_csv) {
    btn_pick_csv.addEventListener('click', () => {
      if (state.ui.is_processing) {
        return;
      }
      if (!state.institution_name) {
        alert('Primero guarda el nombre de la institución.');
        institution_input.focus();
        return;
      }
      open_file_picker();
    });
  }

  ['dragenter', 'dragover'].forEach((event_name) => {
    drop_zone.addEventListener(event_name, (event) => {
      event.preventDefault();
      if (!state.institution_name) {
        return;
      }
      drop_zone.classList.add('is-dragover');
    });
  });

  ['dragleave', 'drop'].forEach((event_name) => {
    drop_zone.addEventListener(event_name, (event) => {
      event.preventDefault();
      drop_zone.classList.remove('is-dragover');
    });
  });

  drop_zone.addEventListener('drop', (event) => {
    if (state.ui.is_processing) {
      return;
    }
    if (!state.institution_name) {
      alert('Primero guarda el nombre de la institución.');
      institution_input.focus();
      return;
    }
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      parse_csv(file);
    }
  });

  filter_assigned.addEventListener('change', () => {
    state.filters.assigned = filter_assigned.value;
    state.table.page = 1;
    render();
  });

  filter_phase.addEventListener('change', () => {
    state.filters.phase = filter_phase.value;
    state.table.page = 1;
    render();
  });

  filter_level.addEventListener('change', () => {
    state.filters.level = filter_level.value;
    state.table.page = 1;
    render();
  });

  filter_time.addEventListener('change', () => {
    state.filters.time = filter_time.value;
    toggle_time_custom_range();
    state.table.page = 1;
    render();
  });

  filter_date_from.addEventListener('change', () => {
    state.filters.date_from = filter_date_from.value || '';
    state.table.page = 1;
    render();
  });

  filter_date_to.addEventListener('change', () => {
    state.filters.date_to = filter_date_to.value || '';
    state.table.page = 1;
    render();
  });

  btn_reset.addEventListener('click', () => {
    state.filters = {
      assigned: 'Todos',
      phase: 'Todas',
      level: 'Todos',
      time: 'Todo',
      date_from: '',
      date_to: ''
    };
    state.table.page = 1;
    render_filters();
    render();
  });

  btn_export_csv.addEventListener('click', () => {
    download_filtered_csv();
  });

  btn_export_report.addEventListener('click', async () => {
    await download_report_capture_png();
  });

  btn_export_pdf.addEventListener('click', async () => {
    await download_report_pdf();
  });

  btn_main_action.addEventListener('click', () => {
    handle_main_action();
  });

  institution_input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      save_institution_name();
    }
  });

  btn_toggle_table.addEventListener('click', () => {
    state.table.is_collapsed = !state.table.is_collapsed;
    render_table(state.filtered_rows);
  });

  btn_page_prev.addEventListener('click', () => {
    if (state.table.page > 1) {
      state.table.page -= 1;
      render_table(state.filtered_rows);
    }
  });

  btn_page_next.addEventListener('click', () => {
    const total_pages = Math.max(1, Math.ceil(state.filtered_rows.length / state.table.page_size));
    if (state.table.page < total_pages) {
      state.table.page += 1;
      render_table(state.filtered_rows);
    }
  });

  inject_chart_info_buttons();
  strip_chart_title_numbers();
}

function inject_chart_info_buttons() {
  const cards = document.querySelectorAll('.chart-card');
  cards.forEach((card) => {
    const head = card.querySelector('.chart-head');
    const title_node = head?.querySelector('h3');
    const desc_node = head?.querySelector('p');
    if (!head || !title_node || !desc_node || head.querySelector('.chart-info-btn')) {
      return;
    }

    const title = clean(title_node.textContent);
    const desc = clean(desc_node.textContent);
    const info_btn = document.createElement('button');
    info_btn.type = 'button';
    info_btn.className = 'chart-info-btn';
    info_btn.textContent = 'i';
    info_btn.setAttribute('aria-label', `Más información de ${title}`);
    info_btn.title = `${title}: ${desc}`;
    info_btn.setAttribute('data-tooltip', `${title}: ${desc}`);
    const tooltip = document.createElement('span');
    tooltip.className = 'chart-info-tooltip';
    tooltip.textContent = `${title}: ${desc}`;
    info_btn.appendChild(tooltip);
    head.appendChild(info_btn);
  });
}

function strip_chart_title_numbers() {
  const titles = document.querySelectorAll('.chart-head h3');
  titles.forEach((node) => {
    node.textContent = clean(node.textContent).replace(/^\d+\.\s*/, '');
  });
}

function parse_csv(file) {
  if (state.ui.is_processing) {
    return;
  }
  if (!state.institution_name) {
    alert('Primero guarda el nombre de la institución.');
    institution_input.focus();
    return;
  }

  const lower_name = (file.name || '').toLowerCase();
  const is_excel = lower_name.endsWith('.xlsx') || lower_name.endsWith('.xls');

  if (is_excel) {
    parse_excel(file);
    return;
  }

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: (result) => {
      process_loaded_rows(result.data, result.meta.fields || Object.keys(result.data?.[0] || {}), 'CSV');
    },
    error: () => {
      alert('Error al leer el archivo CSV.');
    }
  });
}

function parse_excel(file) {
  if (typeof XLSX === 'undefined') {
    alert('No se pudo cargar el convertidor de Excel. Intenta con CSV.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const first_sheet_name = workbook.SheetNames?.[0];
      if (!first_sheet_name) {
        alert('El archivo Excel no contiene hojas con datos.');
        return;
      }
      const worksheet = workbook.Sheets[first_sheet_name];
      const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
      const headers = Object.keys(rows?.[0] || {});
      process_loaded_rows(rows, headers, 'Excel');
    } catch (_error) {
      alert('No fue posible convertir el archivo Excel.');
    }
  };
  reader.onerror = () => {
    alert('Error al leer el archivo Excel.');
  };
  reader.readAsArrayBuffer(file);
}

function process_loaded_rows(data_rows, headers, source_label) {
  if (!data_rows?.length) {
    alert(`No se encontraron filas en el archivo ${source_label}.`);
    return;
  }

  state.raw_rows = data_rows;
  state.headers = headers?.length ? headers : Object.keys(data_rows[0] || {});
  state.col = resolve_columns(state.headers);
  state.rows = enrich_rows(data_rows, state.col);

  if (!state.rows.length) {
    alert('No fue posible leer filas válidas. Revisa el archivo.');
    return;
  }

      state.filters = {
        assigned: 'Todos',
        phase: 'Todas',
        level: 'Todos',
        time: 'Todo',
        date_from: '',
        date_to: ''
      };
  state.table.page = 1;
  state.table.is_collapsed = true;
  start_analysis_sequence(() => {
    controls.hidden = false;
    charts_wrap.hidden = false;
    table_wrap.hidden = false;
    if (records_wrap) {
      records_wrap.hidden = false;
    }
    render_filters();
    render();
    refresh_main_action_button();
    queue_auto_pdf_download_after_upload();
  });
}

function resolve_columns(headers) {
  const normalized = headers.map((header) => ({
    raw: header,
    norm: normalize_text(header)
  }));

  return {
    lead_name: find_first(normalized, ['nombre del cliente potencial', 'nombre cliente']),
    contact_name: find_first(normalized, ['nombre del contacto']),
    phase: find_first(normalized, ['fase']),
    source: find_first(normalized, ['fuente']),
    assigned: find_first(normalized, ['asignado']),
    value: find_first(normalized, ['valor del cliente potencial', 'valor']),
    created_at: find_first(normalized, ['creado el']),
    updated_at: find_first(normalized, ['actualizado el']),
    engagement: find_first(normalized, ['puntuación de compromiso', 'puntuación']),
    level: find_first(normalized, ['nivel educativo de interés', 'nivel educativo', 'nivel academico', 'nivel académico']),
    grade: find_first(normalized, ['grado educativo de interés', 'grado educativo', 'grado escolar']),
    birth_date: find_first(normalized, ['fecha de nacimiento del aspirante', 'fecha de nacimiento']),
    closed_at: find_first(normalized, ['fecha de cierre']),
    school: find_first(normalized, ['escuela de procedencia']),
    admission_type: find_first(normalized, ['tipo de ingreso']),
    status: find_first(normalized, ['estado']),
    days_last_step: find_first(normalized, ['días desde el último paso']),
    days_last_status: find_first(normalized, ['días desde la fecha del último cambio de estado']),
    days_since_updated: find_first(normalized, ['días desde actualizado el'])
  };
}

function find_first(normalized_headers, candidates) {
  for (const candidate of candidates) {
    const target = normalize_text(candidate);
    const match = normalized_headers.find((h) => h.norm.includes(target));
    if (match) {
      return match.raw;
    }
  }
  return null;
}

function enrich_rows(rows, col) {
  return rows
    .map((row) => {
      const created_at = parse_date(row[col.created_at]);
      const updated_at = parse_date(row[col.updated_at]);
      const closed_at = parse_date(row[col.closed_at]);
      const birth_date = parse_date(row[col.birth_date]);
      const age = calc_age(birth_date);

      return {
        raw: row,
        lead_name: clean(row[col.lead_name]),
        contact_name: clean(row[col.contact_name]),
        phase: normalize_phase(clean(row[col.phase])) || 'Sin fase',
        source: clean(row[col.source]) || 'Sin fuente',
        assigned: clean(row[col.assigned]) || 'Sin asesor',
        value: parse_number(row[col.value]) || 0,
        created_at,
        updated_at,
        closed_at,
        engagement: parse_number(row[col.engagement]),
        level: clean(row[col.level]) || 'Sin nivel',
        grade: clean(row[col.grade]) || 'Sin grado',
        birth_date,
        age,
        school: clean(row[col.school]) || 'Sin escuela',
        admission_type: clean(row[col.admission_type]) || 'Sin tipo',
        status: clean(row[col.status]) || 'Sin estado',
        days_last_step: parse_days(row[col.days_last_step]),
        days_last_status: parse_days(row[col.days_last_status]),
        days_since_updated: parse_days(row[col.days_since_updated])
      };
    })
    .filter((row) => row.lead_name || row.contact_name || row.source || row.phase);
}

function render() {
  const rows = apply_filters(state.rows, state.filters);
  state.filtered_rows = rows;
  update_hero_context(rows.length, state.rows.length);
  render_kpis(rows);
  render_insights(rows);
  render_charts(rows);
  sync_charts_section_visibility();
  render_table(rows);
  render_records(rows);
}

function get_chart_card(id) {
  const chart = document.getElementById(id);
  if (!chart) return null;
  return chart.closest('.chart-card');
}

function set_chart_visibility(id, is_visible) {
  const card = get_chart_card(id);
  if (!card) return;
  card.hidden = !is_visible;
}

function sync_charts_section_visibility() {
  const cards = [...document.querySelectorAll('#charts .chart-card')];
  const visible_cards = cards.filter((card) => !card.hidden);
  charts_wrap.hidden = visible_cards.length === 0;
}

function render_filters() {
  const assigned_values = unique_sorted(state.rows.map((row) => row.assigned));
  const phase_values = unique_sorted(state.rows.map((row) => row.phase), phase_sorter);
  const level_values = unique_sorted(state.rows.map((row) => row.level));
  const time_values = [
    'Todo',
    'Hoy',
    'Ayer',
    'Últimos 7 días',
    'Últimos 30 días',
    'Este mes',
    'Mes pasado',
    'Este año',
    'Año pasado',
    'Personalizado'
  ];

  set_select(filter_assigned, ['Todos', ...assigned_values], state.filters.assigned);
  set_select(filter_phase, ['Todas', ...phase_values], state.filters.phase);
  set_select(filter_level, ['Todos', ...level_values], state.filters.level);
  set_select(filter_time, time_values, state.filters.time);
  filter_date_from.value = state.filters.date_from || '';
  filter_date_to.value = state.filters.date_to || '';
  toggle_time_custom_range();
}

function set_select(element, values, current) {
  element.innerHTML = values
    .map((value) => `<option value="${escape_html(value)}" ${value === current ? 'selected' : ''}>${escape_html(value)}</option>`)
    .join('');
}

function save_institution_name() {
  const name = clean(institution_input.value);
  if (!name) {
    alert('Escribe el nombre de la institución para continuar.');
    institution_input.focus();
    refresh_main_action_button();
    return false;
  }
  state.institution_name = name;
  safe_storage.set('dashboard_institution_name', name);
  apply_institution_name(name);
  refresh_main_action_button();
  return true;
}

function open_file_picker() {
  if (!csv_input) {
    return;
  }

  try {
    csv_input.value = '';
  } catch (_error) {
    // Ignorar.
  }

  if (typeof csv_input.showPicker === 'function') {
    try {
      csv_input.showPicker();
      return;
    } catch (_error) {
      // Fallback a click.
    }
  }

  csv_input.click();
}

function handle_main_action() {
  if (state.ui.is_processing) {
    return;
  }
  const ok = save_institution_name();
  if (!ok) {
    return;
  }
  focus_drop_zone();
}

function refresh_main_action_button() {
  if (!btn_main_action) {
    return;
  }
  if (state.ui.is_processing) {
    btn_main_action.textContent = 'Analizando archivo...';
    btn_main_action.classList.remove('is-secondary');
    return;
  }

  const has_saved_name = Boolean(state.institution_name);

  if (!has_saved_name) {
    btn_main_action.textContent = 'Guardar nombre';
    btn_main_action.classList.remove('is-secondary');
    return;
  }

  btn_main_action.textContent = 'Actualizar nombre';
  btn_main_action.classList.add('is-secondary');
}

function apply_institution_name(name) {
  const clean_name = clean(name);
  const base_label = 'Análisis admisiones SuperLeads';
  dashboard_title.textContent = base_label;
  document.title = clean_name ? `${base_label} - ${clean_name}` : base_label;
  institution_hero_box.textContent = clean_name;
  institution_hero_box.hidden = !clean_name;
  hero_subtitle.textContent = clean_name
    ? `No es intuición. Es evidencia operativa para ${clean_name}.`
    : 'No más esfuerzo desordenado. Sí más sistema para inscripciones.';
  sync_drop_zone_state();
}

function clear_analysis_timers() {
  state.ui.analysis_timer_ids.forEach((timer_id) => window.clearTimeout(timer_id));
  state.ui.analysis_timer_ids = [];
}

function reset_analysis_sequence_ui() {
  clear_analysis_timers();
  analysis_steps.forEach((step) => {
    if (!step) return;
    step.classList.remove('is-active', 'is-done');
  });
  if (analysis_progress_bar) {
    analysis_progress_bar.style.width = '0%';
  }
}

function start_analysis_sequence(on_complete) {
  reset_analysis_sequence_ui();
  state.ui.is_processing = true;
  controls.hidden = true;
  charts_wrap.hidden = true;
  table_wrap.hidden = true;
  if (records_wrap) {
    records_wrap.hidden = true;
  }
  analysis_sequence.hidden = false;
  drop_zone.classList.add('is-disabled');
  btn_main_action.disabled = true;
  refresh_main_action_button();

  const phases = [
    {
      title: 'Analizando datos',
      message: 'Estamos validando la estructura del archivo CSV.',
      progress: 20
    },
    {
      title: 'Acomodando información',
      message: 'Estamos normalizando columnas y limpiando registros.',
      progress: 40
    },
    {
      title: 'Graficando indicadores',
      message: 'Estamos construyendo las visualizaciones principales.',
      progress: 62
    },
    {
      title: 'Detallando hallazgos',
      message: 'Estamos calculando métricas, filtros y tabla de detalle.',
      progress: 84
    },
    {
      title: 'Exponiendo resultados',
      message: 'Estamos preparando el tablero final para presentación.',
      progress: 100
    }
  ];

  let delay = 0;
  phases.forEach((phase, index) => {
    const timer_id = window.setTimeout(() => {
      if (analysis_title) analysis_title.textContent = phase.title;
      if (analysis_message) analysis_message.textContent = phase.message;
      if (analysis_progress_bar) analysis_progress_bar.style.width = `${phase.progress}%`;

      analysis_steps.forEach((step, step_index) => {
        if (!step) return;
        step.classList.toggle('is-active', step_index === index);
        step.classList.toggle('is-done', step_index < index);
      });
    }, delay);
    state.ui.analysis_timer_ids.push(timer_id);
    delay += 780;
  });

  const finish_timer_id = window.setTimeout(() => {
    analysis_steps.forEach((step) => step?.classList.remove('is-active'));
    analysis_sequence.hidden = true;
    state.ui.is_processing = false;
    btn_main_action.disabled = false;
    refresh_main_action_button();
    sync_drop_zone_state();
    on_complete();
  }, delay + 260);
  state.ui.analysis_timer_ids.push(finish_timer_id);
}

function sync_drop_zone_state() {
  const should_disable = !state.institution_name || state.ui.is_processing;
  drop_zone.classList.toggle('is-disabled', should_disable);
}

function focus_drop_zone() {
  drop_zone.classList.add('is-focus');
  drop_zone.scrollIntoView({ behavior: 'smooth', block: 'center' });
  drop_zone.focus();
  window.setTimeout(() => drop_zone.classList.remove('is-focus'), 1100);
}

function update_hero_context(filtered, total) {
  if (!state.institution_name || !total) {
    return;
  }
  const suffix = filtered === total
    ? `${fmt_int(total)} registros activos`
    : `${fmt_int(filtered)} de ${fmt_int(total)} registros filtrados`;
  hero_subtitle.textContent = `Más control. Más claridad. Más inscripciones para ${state.institution_name}. ${suffix}.`;
}

function queue_auto_pdf_download_after_upload() {
  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(async () => {
      try {
        await download_report_pdf();
      } catch (error) {
        console.error(error);
      }
    });
  });
}

async function build_dashboard_canvas() {
  if (!state.rows.length) {
    alert('Carga un CSV antes de descargar el reporte.');
    return null;
  }

  if (typeof html2canvas !== 'function') {
    alert('No se pudo cargar la librería de captura (html2canvas).');
    return null;
  }

  const target = document.getElementById('sladm-root');
  return html2canvas(target, {
    backgroundColor: '#001240',
    scale: 2.4,
    useCORS: true
  });
}

async function download_report_capture_png() {
  const original_label = btn_export_report.textContent;
  btn_export_report.disabled = true;
  btn_export_report.textContent = 'Generando imagen...';

  try {
    const canvas = await build_dashboard_canvas();
    if (!canvas) {
      return;
    }
    const file_name = build_report_file_name('png');
    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/png');
    link.download = file_name;
    link.click();
  } catch (error) {
    console.error(error);
    alert('No se pudo generar la imagen del reporte.');
  } finally {
    btn_export_report.disabled = false;
    btn_export_report.textContent = original_label;
  }
}

async function download_report_pdf() {
  const original_label = btn_export_pdf.textContent;
  btn_export_pdf.disabled = true;
  btn_export_pdf.textContent = 'Generando PDF...';

  try {
    const canvas = await build_dashboard_canvas();
    if (!canvas) {
      return;
    }

    const js_pdf = window.jspdf?.jsPDF;
    if (typeof js_pdf !== 'function') {
      alert('No se pudo cargar la librería PDF (jsPDF).');
      return;
    }

    const pdf = new js_pdf({
      orientation: 'p',
      unit: 'mm',
      format: 'a4',
      compress: true
    });
    const page_width = 210;
    const page_height = 297;
    const px_per_mm = canvas.width / page_width;
    const page_height_px = Math.floor(page_height * px_per_mm);
    let offset_px = 0;
    let page_index = 0;

    while (offset_px < canvas.height) {
      const slice_height_px = Math.min(page_height_px, canvas.height - offset_px);
      const slice_canvas = document.createElement('canvas');
      slice_canvas.width = canvas.width;
      slice_canvas.height = slice_height_px;
      const slice_ctx = slice_canvas.getContext('2d');
      if (!slice_ctx) {
        throw new Error('No fue posible crear el contexto del PDF.');
      }

      slice_ctx.fillStyle = '#001240';
      slice_ctx.fillRect(0, 0, slice_canvas.width, slice_canvas.height);
      slice_ctx.drawImage(
        canvas,
        0, offset_px, canvas.width, slice_height_px,
        0, 0, slice_canvas.width, slice_canvas.height
      );

      const img_data = slice_canvas.toDataURL('image/png');
      const slice_height_mm = (slice_height_px * page_width) / canvas.width;
      if (page_index > 0) {
        pdf.addPage();
      }

      pdf.setFillColor(0, 18, 64);
      pdf.rect(0, 0, page_width, page_height, 'F');
      pdf.addImage(img_data, 'PNG', 0, 0, page_width, slice_height_mm, undefined, 'FAST');

      offset_px += slice_height_px;
      page_index += 1;
    }

    pdf.save(build_report_file_name('pdf'));
  } catch (error) {
    console.error(error);
    alert('No se pudo generar el PDF del reporte.');
  } finally {
    btn_export_pdf.disabled = false;
    btn_export_pdf.textContent = original_label;
  }
}

function build_report_file_name(extension) {
  const institution = sanitize_institution_for_file(state.institution_name || 'Colegio');
  const stamp = file_stamp_full(new Date());
  return `SuperLeads_${institution}_Admisiones_${stamp}.${extension}`;
}

function download_filtered_csv() {
  if (!state.filtered_rows.length) {
    alert('No hay registros para descargar en CSV.');
    return;
  }

  if (!state.headers.length) {
    alert('No se detectaron columnas para exportar.');
    return;
  }

  const rows = state.filtered_rows.map((row) => row.raw || {});
  const csv_text = Papa.unparse({
    fields: state.headers,
    data: rows.map((row) => state.headers.map((header) => row[header] ?? ''))
  });

  const blob = new Blob([csv_text], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = build_report_file_name('csv');
  link.click();
  URL.revokeObjectURL(url);
}

function apply_filters(rows, filters) {
  const time_matcher = build_time_matcher(filters);
  return rows.filter((row) => {
    const ok_assigned = filters.assigned === 'Todos' || row.assigned === filters.assigned;
    const ok_phase = filters.phase === 'Todas' || row.phase === filters.phase;
    const ok_level = filters.level === 'Todos' || row.level === filters.level;
    const ok_time = time_matcher(get_row_reference_date(row));
    return ok_assigned && ok_phase && ok_level && ok_time;
  });
}

function toggle_time_custom_range() {
  if (!time_custom_range) {
    return;
  }
  time_custom_range.hidden = state.filters.time !== 'Personalizado';
}

function get_row_reference_date(row) {
  return row.created_at || row.updated_at || row.closed_at || null;
}

function build_time_matcher(filters) {
  const now = new Date();
  const start_of_today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const end_of_today = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  const start_of_this_month = new Date(now.getFullYear(), now.getMonth(), 1);
  const start_of_next_month = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const start_of_prev_month = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const start_of_this_year = new Date(now.getFullYear(), 0, 1);
  const start_of_next_year = new Date(now.getFullYear() + 1, 0, 1);
  const start_of_prev_year = new Date(now.getFullYear() - 1, 0, 1);

  function between(date, from, to) {
    return date >= from && date <= to;
  }

  if (filters.time === 'Todo') {
    return () => true;
  }

  if (filters.time === 'Hoy') {
    return (date) => date instanceof Date && between(date, start_of_today, end_of_today);
  }

  if (filters.time === 'Ayer') {
    const start = new Date(start_of_today);
    start.setDate(start.getDate() - 1);
    const end = new Date(start_of_today);
    end.setMilliseconds(-1);
    return (date) => date instanceof Date && between(date, start, end);
  }

  if (filters.time === 'Últimos 7 días') {
    const start = new Date(start_of_today);
    start.setDate(start.getDate() - 6);
    return (date) => date instanceof Date && between(date, start, end_of_today);
  }

  if (filters.time === 'Últimos 30 días') {
    const start = new Date(start_of_today);
    start.setDate(start.getDate() - 29);
    return (date) => date instanceof Date && between(date, start, end_of_today);
  }

  if (filters.time === 'Este mes') {
    return (date) => date instanceof Date && date >= start_of_this_month && date < start_of_next_month;
  }

  if (filters.time === 'Mes pasado') {
    return (date) => date instanceof Date && date >= start_of_prev_month && date < start_of_this_month;
  }

  if (filters.time === 'Este año') {
    return (date) => date instanceof Date && date >= start_of_this_year && date < start_of_next_year;
  }

  if (filters.time === 'Año pasado') {
    return (date) => date instanceof Date && date >= start_of_prev_year && date < start_of_this_year;
  }

  if (filters.time === 'Personalizado') {
    const from = filters.date_from ? new Date(`${filters.date_from}T00:00:00`) : null;
    const to = filters.date_to ? new Date(`${filters.date_to}T23:59:59.999`) : null;
    return (date) => {
      if (!(date instanceof Date)) return false;
      if (from && date < from) return false;
      if (to && date > to) return false;
      return true;
    };
  }

  return () => true;
}

function render_kpis(rows) {
  const total = rows.length;
  const citas = rows.filter((row) => normalize_text(row.phase).includes('citas')).length;
  const contactados = rows.filter((row) => {
    const p = normalize_text(row.phase);
    return p.includes('contactado') || p.includes('citas') || p.includes('proceso');
  }).length;
  const avg_engagement = average(rows.map((row) => row.engagement));
  const avg_days_update = average(rows.map((row) => row.days_since_updated));
  const avg_age = average(rows.map((row) => row.age));

  const conversion_citas = total > 0 ? (citas / total) * 100 : 0;

  const cards = [
    ['Total oportunidades', fmt_int(total)],
    ['Contactados/proceso', fmt_int(contactados)],
    ['Citas', fmt_int(citas)],
    ['Conversión a citas', `${conversion_citas.toFixed(1)}%`],
    ['Compromiso promedio', fmt_number(avg_engagement, 1)],
    ['Días desde actualización', fmt_number(avg_days_update, 1)],
    ['Edad promedio aspirante', Number.isFinite(avg_age) ? `${avg_age.toFixed(1)} años` : 'N/D']
  ];

  kpi_grid.innerHTML = cards
    .map(([title, value]) => `<article class="kpi"><h3>${title}</h3><p>${value}</p></article>`)
    .join('');
}

function render_insights(rows) {
  if (!rows.length) {
    insights_box.innerHTML = '<h2>Hallazgos rápidos</h2><p class="empty-note">Sin datos para mostrar con los filtros actuales.</p>';
    return;
  }

  const top_source = top_item(count_by(rows, (row) => row.source));
  const top_phase = top_item(count_by(rows, (row) => row.phase));
  const top_assigned = top_item(count_by(rows, (row) => row.assigned));

  const old_rows = rows.filter((row) => Number.isFinite(row.days_since_updated) && row.days_since_updated >= 3).length;
  const stale_ratio = (old_rows / rows.length) * 100;
  const citas = rows.filter((row) => row.phase === 'Citas').length;
  const cita_rate = rows.length ? (citas / rows.length) * 100 : 0;

  const notes = [
    `Dato: ${top_source.label} concentra ${fmt_int(top_source.value)} registros | Significado: este canal es tu principal generador de demanda y merece prioridad de inversión.`,
    `Dato: ${top_phase.label} tiene ${fmt_int(top_phase.value)} oportunidades | Significado: aquí está el mayor volumen operativo del pipeline actual.`,
    `Dato: ${top_assigned.label} gestiona ${fmt_int(top_assigned.value)} casos | Significado: esta persona concentra mayor carga y puede requerir balance de distribución.`,
    `Dato: ${cita_rate.toFixed(1)}% del universo filtrado está en Citas | Significado: esta tasa refleja el avance real hacia cierre y calidad de seguimiento.`,
    `Dato: ${stale_ratio.toFixed(1)}% de registros tiene 3+ días sin actualizar | Significado: existe riesgo de fuga por falta de seguimiento oportuno.`
  ];

  insights_box.innerHTML = `<h2>Hallazgos rápidos</h2><ul>${notes.map((n) => `<li>${n}</li>`).join('')}</ul>`;
}

function render_charts(rows) {
  render_funnel(rows);
  render_created_trend(rows);
  render_won_students(rows);
  render_phase_pie(rows);
  render_source_bar(rows);
  render_source_conversion(rows);
  render_assigned_bar(rows);
  render_level_grade(rows);
  render_heatmap(rows);
  render_heatmap_day(rows);
  render_heatmap_month(rows);
  render_age_hist(rows);
  render_days_box(rows);
}

function render_funnel(rows) {
  const counts = count_by(rows, (row) => row.phase);
  const labels = Object.keys(counts)
    .filter((label) => !is_excluded_funnel_phase(label))
    .sort(phase_sorter);
  const values = labels.map((label) => counts[label]);
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!values.length) {
    render_empty_chart('chart-funnel', 'Embudo por fase', 'Sin datos para construir el embudo.');
    return;
  }

  const pct = values.map((value) => (total ? (value / total) * 100 : 0));
  const text_labels = values.map((value, index) => `${fmt_int(value)} (${pct[index].toFixed(1)}%)`);

  plot_chart(
    'chart-funnel',
    [{
      type: 'funnel',
      y: labels,
      x: values,
      customdata: pct,
      text: text_labels,
      textposition: 'inside',
      textfont: { size: 13, color: '#001240' },
      marker: { color: labels.map((label) => phase_palette[label] || chart_colors.primary) },
      hovertemplate: '<b>Etapa:</b> %{y}<br><b>Personas:</b> %{x}<br><b>% del total filtrado:</b> %{customdata:.1f}%<extra></extra>'
    }],
    base_layout('Embudo por fase')
  );
}

function is_excluded_funnel_phase(label) {
  const text = normalize_text(label);
  return (
    text.includes('perdid') ||
    text.includes('abandon') ||
    text.includes('descartado') ||
    text.includes('cancelado')
  );
}

function render_created_trend(rows) {
  const counts = count_by(
    rows.filter((row) => row.created_at),
    (row) => to_local_date_key(row.created_at)
  );
  const labels = Object.keys(counts).sort();
  const values = labels.map((label) => counts[label]);
  const total = values.reduce((sum, value) => sum + value, 0);
  const pct = values.map((value) => (total ? (value / total) * 100 : 0));
  if (!values.length) {
    render_empty_chart('chart-created-trend', 'Captación de leads en admisiones', 'No hay fechas de alta para graficar.');
    return;
  }

  plot_chart(
    'chart-created-trend',
    [{
      type: 'scatter',
      mode: 'lines+markers',
      x: labels,
      y: values,
      customdata: pct,
      line: { color: chart_colors.secondary, width: 3, shape: 'spline' },
      marker: { size: 7, color: '#FFFFFF', line: { width: 2, color: chart_colors.secondary } },
      fill: 'tozeroy',
      fillcolor: 'rgba(42,137,251,0.24)',
      hovertemplate: '<b>Fecha:</b> %{x}<br><b>Nuevos registros:</b> %{y}<br><b>% del total del periodo:</b> %{customdata:.1f}%<extra></extra>'
    }],
    base_layout('Captación de leads en admisiones')
  );
}

function render_won_students(rows) {
  const strategy = resolve_won_phase_strategy(rows);
  const won_rows = rows.filter((row) => strategy.match(normalize_phase(row.phase)));
  const counts = count_by(
    won_rows.filter((row) => row.closed_at || row.updated_at || row.created_at),
    (row) => to_local_date_key(row.closed_at || row.updated_at || row.created_at)
  );
  const labels = Object.keys(counts).sort();
  const values = labels.map((label) => counts[label]);
  const total = values.reduce((sum, value) => sum + value, 0);
  const pct = values.map((value) => (total ? (value / total) * 100 : 0));

  if (!values.length) {
    render_empty_chart(
      'chart-won-students',
      strategy.chart_title,
      strategy.empty_message
    );
    return;
  }

  plot_chart(
    'chart-won-students',
    [
      {
        type: 'bar',
        x: labels,
        y: values,
        customdata: pct,
        marker: { color: '#56EF9F', line: { width: 1, color: '#2BC878' } },
        hovertemplate: `<b>Fecha:</b> %{x}<br><b>Nuevos alumnos (${strategy.hover_label}):</b> %{y}<br><b>% del total de nuevos alumnos:</b> %{customdata:.1f}%<extra></extra>`
      },
      {
        type: 'scatter',
        mode: 'lines+markers',
        x: labels,
        y: values,
        customdata: pct,
        line: { color: '#2A89FB', width: 2, shape: 'spline' },
        marker: { size: 6, color: '#FFFFFF', line: { width: 2, color: '#2A89FB' } },
        hovertemplate: `<b>Fecha:</b> %{x}<br><b>Tendencia (${strategy.hover_label}):</b> %{y}<br><b>% del total de nuevos alumnos:</b> %{customdata:.1f}%<extra></extra>`
      }
    ],
    base_layout(strategy.chart_title)
  );
}

function render_phase_pie(rows) {
  const counts = count_by(rows, (row) => row.phase);
  const labels = Object.keys(counts)
    .filter((label) => !is_excluded_funnel_phase(label))
    .sort(phase_sorter);
  const values = labels.map((label) => counts[label]);
  const total = values.reduce((sum, value) => sum + value, 0);
  const pct = values.map((value) => (total ? (value / total) * 100 : 0));
  if (!values.length) {
    render_empty_chart('chart-phase', 'Distribución por fase', 'No hay fases disponibles en este filtro.');
    return;
  }

  plot_chart(
    'chart-phase',
    [{
      type: 'pie',
      labels,
      values,
      customdata: pct,
      hole: 0.5,
      textinfo: 'percent',
      textposition: 'inside',
      marker: { colors: labels.map((label) => phase_palette[label] || '#EEF2FF') },
      hovertemplate: '<b>Fase:</b> %{label}<br><b>Personas:</b> %{value}<br><b>% del total filtrado:</b> %{customdata:.1f}%<extra></extra>',
      sort: false
    }],
    {
      ...base_layout('Distribución por fase'),
      showlegend: true,
      legend: { orientation: 'h', y: -0.16 }
    }
  );
}

function render_source_bar(rows) {
  const counts = Object.entries(limit_map(count_by(rows, (row) => row.source), 10)).sort((a, b) => b[1] - a[1]);
  const total = counts.reduce((sum, entry) => sum + entry[1], 0);
  const labels = counts.map((entry) => truncate_label(entry[0], 36)).reverse();
  const values = counts.map((entry) => entry[1]).reverse();
  const pct = values.map((value) => (total ? (value / total) * 100 : 0));
  if (!values.length) {
    render_empty_chart('chart-source', 'Volumen por fuente', 'No hay fuentes para comparar.');
    return;
  }

  plot_chart(
    'chart-source',
    [{
      type: 'bar',
      x: values,
      y: labels,
      customdata: pct,
      orientation: 'h',
      marker: { color: chart_colors.primary, line: { width: 1, color: '#002582' } },
      hovertemplate: '<b>Fuente:</b> %{y}<br><b>Personas:</b> %{x}<br><b>% del total en top fuentes:</b> %{customdata:.1f}%<extra></extra>'
    }],
    base_layout('Volumen por fuente')
  );
}

function render_source_conversion(rows) {
  const grouped = {};
  rows.forEach((row) => {
    const source = row.source || 'Sin fuente';
    if (!grouped[source]) {
      grouped[source] = { total: 0, citas: 0 };
    }
    grouped[source].total += 1;
    if (is_cita_or_later(row.phase)) {
      grouped[source].citas += 1;
    }
  });

  const ranked = Object.entries(grouped)
    .map(([source, value]) => ({
      source,
      total: value.total,
      citas: value.citas,
      rate: value.total ? (value.citas / value.total) * 100 : 0
    }))
    .sort((a, b) => b.rate - a.rate || b.total - a.total)
    .slice(0, 10);

  if (!ranked.length) {
    render_empty_chart(
      'chart-source-conversion',
      'Conversión a cita por fuente',
      'No hay fuentes con datos para calcular conversión.'
    );
    return;
  }

  const labels = ranked.map((item) => truncate_label(item.source, 34)).reverse();
  const rates = ranked.map((item) => item.rate).reverse();
  const totals = ranked.map((item) => item.total).reverse();
  const citas = ranked.map((item) => item.citas).reverse();

  const axis_base = base_layout('');
  plot_chart(
    'chart-source-conversion',
    [{
      type: 'bar',
      x: rates,
      y: labels,
      orientation: 'h',
      marker: {
        color: rates,
        colorscale: [
          [0, '#EEF2FF'],
          [0.5, '#2A89FB'],
          [1, '#0039C8']
        ],
        line: { color: '#002582', width: 1 }
      },
      customdata: totals.map((total, index) => [total, citas[index]]),
      text: rates.map((rate) => `${rate.toFixed(1)}%`),
      textposition: 'outside',
      cliponaxis: false,
      hovertemplate: '<b>Fuente:</b> %{y}<br><b>Conversión a cita:</b> %{x:.1f}%<br><b>Total evaluado:</b> %{customdata[0]}<br><b>Citas o etapas posteriores:</b> %{customdata[1]}<extra></extra>'
    }],
    {
      ...base_layout('Conversión a cita por fuente'),
      xaxis: { ...axis_base.xaxis, range: [0, 100], ticksuffix: '%', title: { text: 'Tasa de conversion' } }
    }
  );
}

function render_assigned_bar(rows) {
  const counts = Object.entries(count_by(rows, (row) => row.assigned))
    .filter((entry) => Number(entry[1]) > 0 && clean(entry[0]))
    .sort((a, b) => b[1] - a[1]);
  const total = counts.reduce((sum, entry) => sum + entry[1], 0);
  const labels = counts.map((entry) => entry[0]);
  const values = counts.map((entry) => entry[1]);
  const pct = values.map((value) => (total ? (value / total) * 100 : 0));
  if (!values.length) {
    render_empty_chart('chart-assigned', 'Carga por asesor', 'No hay asesores para mostrar.');
    return;
  }

  plot_chart(
    'chart-assigned',
    [{
      type: 'bar',
      x: labels,
      y: values,
      customdata: pct,
      marker: {
        color: values,
        colorscale: [
          [0, '#EEF2FF'],
          [0.45, '#1DB2FC'],
          [1, '#0039C8']
        ],
        showscale: false,
        line: { color: '#0039C8', width: 1 }
      },
      hovertemplate: '<b>Asesor:</b> %{x}<br><b>Registros:</b> %{y}<br><b>% del total filtrado:</b> %{customdata:.1f}%<extra></extra>'
    }],
    base_layout('Carga por asesor', -20)
  );
}

function render_level_grade(rows) {
  const counts = Object.entries(count_by(rows, (row) => `${row.level} | ${row.grade}`))
    .sort((a, b) => b[1] - a[1])
    .slice(0, 14);
  const total = counts.reduce((sum, entry) => sum + entry[1], 0);
  const labels = counts.map((entry) => truncate_label(entry[0], 28));
  const values = counts.map((entry) => entry[1]);
  const pct = values.map((value) => (total ? (value / total) * 100 : 0));
  if (!values.length) {
    render_empty_chart('chart-level-grade', 'Combinaciones nivel y grado', 'No hay niveles o grados para visualizar.');
    return;
  }

  plot_chart(
    'chart-level-grade',
    [{
      type: 'bar',
      x: labels,
      y: values,
      customdata: pct,
      marker: { color: chart_colors.secondary, line: { color: '#0039C8', width: 1 } },
      hovertemplate: '<b>Nivel | Grado:</b> %{x}<br><b>Registros:</b> %{y}<br><b>% del total en top combinaciones:</b> %{customdata:.1f}%<extra></extra>'
    }],
    base_layout('Top combinaciones nivel y grado', -28)
  );
}

function render_heatmap(rows) {
  const days = ['Lun', 'Mar', 'Mie', 'Jue', 'Vie', 'Sab', 'Dom'];
  const hours = [...Array(24).keys()];
  const matrix = days.map(() => hours.map(() => 0));

  rows.forEach((row) => {
    if (!row.created_at) {
      return;
    }
    const day_idx = weekday_mon_index(row.created_at);
    const hour = row.created_at.getHours();
    matrix[day_idx][hour] += 1;
  });

  const total = matrix.flat().reduce((sum, value) => sum + value, 0);
  if (!total) {
    render_empty_chart('chart-heatmap', 'Mapa de calor: día y hora de alta', 'No hay marcas de tiempo de alta para el mapa.');
    return;
  }

  plot_chart(
    'chart-heatmap',
    [{
      type: 'heatmap',
      x: hours,
      y: days,
      z: matrix,
      customdata: matrix.map((row) => row.map((value) => (total ? (value / total) * 100 : 0))),
      colorscale: [
        [0, '#EEF2FF'],
        [0.5, '#2A89FB'],
        [1, '#0039C8']
      ],
      hovertemplate: '<b>Día/Hora:</b> %{y} - %{x}:00<br><b>Altas:</b> %{z}<br><b>% de altas con fecha:</b> %{customdata:.1f}%<extra></extra>'
    }],
    base_layout('Mapa de calor: día y hora de alta')
  );
}

function render_heatmap_day(rows) {
  const dated_rows = rows.filter((row) => row.created_at instanceof Date);
  if (!dated_rows.length) {
    render_empty_chart('chart-heatmap-day', 'Mapa de calor: día de alta', 'No hay fechas de alta para este mapa.');
    return;
  }

  const month_keys = unique_sorted(
    dated_rows.map((row) => `${row.created_at.getFullYear()}-${String(row.created_at.getMonth() + 1).padStart(2, '0')}`)
  );
  const month_labels = month_keys.map((key) => {
    const [year_text, month_text] = key.split('-');
    const date = new Date(Number(year_text), Number(month_text) - 1, 1);
    const month_name = date.toLocaleDateString('es-MX', { month: 'short' }).replace('.', '');
    return `${month_name} ${year_text}`;
  });
  const day_numbers = Array.from({ length: 31 }, (_, i) => i + 1);

  const matrix = month_keys.map((key) => {
    const [year_text, month_text] = key.split('-');
    const y = Number(year_text);
    const m = Number(month_text) - 1;
    return day_numbers.map((day) =>
      dated_rows.filter((row) => row.created_at.getFullYear() === y && row.created_at.getMonth() === m && row.created_at.getDate() === day).length
    );
  });

  const total = matrix.flat().reduce((sum, value) => sum + value, 0);
  const pct_matrix = matrix.map((line) => line.map((value) => (total ? (value / total) * 100 : 0)));

  plot_chart(
    'chart-heatmap-day',
    [{
      type: 'heatmap',
      x: day_numbers,
      y: month_labels,
      z: matrix,
      customdata: pct_matrix,
      colorscale: [
        [0, '#EEF2FF'],
        [0.5, '#2A89FB'],
        [1, '#0039C8']
      ],
      hovertemplate: '<b>Mes:</b> %{y}<br><b>Día:</b> %{x}<br><b>Altas:</b> %{z}<br><b>% sobre altas con fecha:</b> %{customdata:.1f}%<extra></extra>'
    }],
    {
      ...base_layout('Mapa de calor: día de alta'),
      xaxis: { ...base_layout('').xaxis, title: { text: 'Día del mes' } }
    }
  );
}

function render_heatmap_month(rows) {
  const dated_rows = rows.filter((row) => row.created_at instanceof Date);
  if (!dated_rows.length) {
    render_empty_chart('chart-heatmap-month', 'Mapa de calor: mes de alta', 'No hay fechas de alta para este mapa.');
    return;
  }

  const years = unique_sorted(dated_rows.map((row) => String(row.created_at.getFullYear())));
  const month_labels = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const matrix = month_labels.map((_month_label, month_index) =>
    years.map((year_text) =>
      dated_rows.filter((row) => String(row.created_at.getFullYear()) === year_text && row.created_at.getMonth() === month_index).length
    )
  );
  const total = matrix.flat().reduce((sum, value) => sum + value, 0);
  const pct_matrix = matrix.map((line) => line.map((value) => (total ? (value / total) * 100 : 0)));

  plot_chart(
    'chart-heatmap-month',
    [{
      type: 'heatmap',
      x: years,
      y: month_labels,
      z: matrix,
      customdata: pct_matrix,
      colorscale: [
        [0, '#EEF2FF'],
        [0.5, '#2A89FB'],
        [1, '#0039C8']
      ],
      hovertemplate: '<b>Año:</b> %{x}<br><b>Mes:</b> %{y}<br><b>Altas:</b> %{z}<br><b>% sobre altas con fecha:</b> %{customdata:.1f}%<extra></extra>'
    }],
    {
      ...base_layout('Mapa de calor: mes de alta'),
      xaxis: { ...base_layout('').xaxis, title: { text: 'Año' } }
    }
  );
}

function render_age_hist(rows) {
  const ages = rows.map((row) => row.age).filter((value) => Number.isFinite(value));
  const total = ages.length;
  if (!ages.length) {
    render_empty_chart('chart-age', 'Distribución de edades de aspirantes', 'No hay fecha de nacimiento suficiente para edades.');
    return;
  }

  plot_chart(
    'chart-age',
    [{
      type: 'histogram',
      x: ages,
      histnorm: '',
      meta: total,
      nbinsx: 12,
      marker: { color: chart_colors.accent, line: { width: 1, color: '#0039C8' } },
      hovertemplate: '<b>Rango de edad:</b> %{x}<br><b>Personas en rango:</b> %{y}<br><b>Base con edad válida:</b> %{meta}<extra></extra>'
    }],
    base_layout('Distribución de edades de aspirantes')
  );
}

function render_days_box(rows) {
  const groups = count_by(rows, (row) => row.phase);
  const labels = Object.keys(groups).sort(phase_sorter);

  const traces = labels.map((phase) => ({
    type: 'box',
    name: phase,
    marker: { color: phase_palette[phase] || chart_colors.neutral },
    line: { color: phase_palette[phase] || chart_colors.neutral },
    hovertemplate: '<b>Fase:</b> %{fullData.name}<br><b>Días sin actualizar:</b> %{y}<extra></extra>',
    y: rows
      .filter((row) => row.phase === phase)
      .map((row) => row.days_since_updated)
      .filter((value) => Number.isFinite(value))
  }));
  if (!traces.length || !traces.some((trace) => trace.y.length)) {
    render_empty_chart('chart-box-days', 'Días sin actualizar por fase', 'No hay datos de seguimiento para boxplot.');
    return;
  }

  const axis_base = base_layout('');
  const layout_base = base_layout('Días sin actualizar por fase', -22);
  plot_chart(
    'chart-box-days',
    traces,
    {
      ...layout_base,
      showlegend: false,
      margin: { ...layout_base.margin, b: 108, r: 16 },
      xaxis: {
        ...axis_base.xaxis,
        tickangle: -22,
        tickfont: { size: 11 },
        automargin: true,
        categoryorder: 'array',
        categoryarray: labels
      },
      yaxis: { ...axis_base.yaxis, title: { text: 'Días' } }
    }
  );
}

function render_table(rows) {
  const table = document.getElementById('data-table');
  const headers = [
    'Alumno',
    'Contacto',
    'Fase',
    'Fuente',
    'Asesor',
    'Nivel',
    'Grado',
    'Compromiso',
    'Días sin actualizar',
    'Creado'
  ];

  const sorted = [...rows].sort((a, b) => (b.created_at?.getTime() || 0) - (a.created_at?.getTime() || 0));
  const total_pages = Math.max(1, Math.ceil(sorted.length / state.table.page_size));
  if (state.table.page > total_pages) {
    state.table.page = total_pages;
  }
  const start = (state.table.page - 1) * state.table.page_size;
  const end = start + state.table.page_size;
  const paginated_rows = sorted.slice(start, end);

  const body = paginated_rows
    .map((row) => {
      return `<tr>
        <td>${escape_html(row.lead_name || '-')}</td>
        <td>${escape_html(row.contact_name || '-')}</td>
        <td>${escape_html(row.phase || '-')}</td>
        <td>${escape_html(row.source || '-')}</td>
        <td>${escape_html(row.assigned || '-')}</td>
        <td>${escape_html(row.level || '-')}</td>
        <td>${escape_html(row.grade || '-')}</td>
        <td>${Number.isFinite(row.engagement) ? row.engagement : '-'}</td>
        <td>${Number.isFinite(row.days_since_updated) ? row.days_since_updated : '-'}</td>
        <td>${row.created_at ? format_date(row.created_at) : '-'}</td>
      </tr>`;
    })
    .join('');

  table.innerHTML = `<thead><tr>${headers.map((h) => `<th>${h}</th>`).join('')}</tr></thead><tbody>${body}</tbody>`;
  table_scroll.hidden = state.table.is_collapsed;
  table_pagination.hidden = state.table.is_collapsed;
  btn_toggle_table.textContent = state.table.is_collapsed ? 'Ver detalle' : 'Ocultar detalle';
  table_page_info.textContent = `Página ${state.table.page} de ${total_pages} (${fmt_int(sorted.length)} registros)`;
  btn_page_prev.disabled = state.table.page <= 1;
  btn_page_next.disabled = state.table.page >= total_pages;
}

function base_layout(title, x_tick_angle = 0) {
  const is_mobile = window.matchMedia('(max-width: 760px)').matches;
  const is_tablet = !is_mobile && window.matchMedia('(max-width: 1024px)').matches;
  const title_size = is_mobile ? 13 : is_tablet ? 14 : 15;
  const font_size = is_mobile ? 11 : 12;
  const margins = is_mobile
    ? { l: 44, r: 12, t: 44, b: 56 }
    : is_tablet
      ? { l: 50, r: 18, t: 48, b: 58 }
      : { l: 56, r: 24, t: 52, b: 62 };

  return {
    title: { text: title, x: 0.02, xanchor: 'left', font: { size: title_size } },
    margin: margins,
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Roboto, Arial, sans-serif', size: font_size, color: '#FFFFFF' },
    colorway: ['#0039C8', '#2A89FB', '#1DB2FC', '#56EF9F', '#9FB7E8', '#EEF2FF'],
    xaxis: {
      tickangle: x_tick_angle,
      showgrid: true,
      gridcolor: 'rgba(255,255,255,0.14)',
      zeroline: false,
      automargin: true,
      title: { standoff: 6 }
    },
    yaxis: {
      automargin: true,
      showgrid: true,
      gridcolor: 'rgba(255,255,255,0.14)',
      zeroline: false
    },
    hoverlabel: {
      bgcolor: '#001240',
      bordercolor: '#0039C8',
      font: { color: '#FFFFFF', size: is_mobile ? 11 : 12 }
    }
  };
}

function plot_chart(id, traces, layout) {
  set_chart_visibility(id, true);
  Plotly.react(id, traces, layout, {
    displayModeBar: false,
    responsive: true,
    locale: 'es'
  });
}

function render_empty_chart(id, title, message) {
  void title;
  void message;
  set_chart_visibility(id, false);
  if (window.Plotly && typeof window.Plotly.purge === 'function') {
    window.Plotly.purge(id);
  }
}

function normalize_text(value) {
  return (value || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function slugify_file_part(value) {
  return normalize_text(value)
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80) || 'institucion';
}

function sanitize_institution_for_file(value) {
  const raw = (value || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
  if (!raw) {
    return 'Colegio';
  }
  const words = raw
    .split(/\s+/)
    .map((word) => word.replace(/[^A-Za-z0-9]/g, ''))
    .filter(Boolean)
    .map((word) => `${word.charAt(0).toUpperCase()}${word.slice(1).toLowerCase()}`);
  return words.length ? words.join('_').slice(0, 80) : 'Colegio';
}

function clean(value) {
  return (value || '').toString().trim();
}

function normalize_phase(phase) {
  const text = normalize_text(phase);
  if (!text) {
    return '';
  }
  if (text.includes('nueva oportunidad')) {
    return 'Nueva Oportunidad';
  }
  if (text.includes('registro')) {
    return 'Nueva Oportunidad';
  }
  if (text.includes('prospecto')) {
    return 'En Proceso de Contactación';
  }
  if (text.includes('proceso de contact')) {
    return 'En Proceso de Contactación';
  }
  if (text.includes('contactado')) {
    return 'Contactado';
  }
  if (text.includes('recorrido')) {
    return 'Citas';
  }
  if (text.includes('citas') || text.includes('cita')) {
    return 'Citas';
  }
  if (text.includes('pendiente firma') || text.includes('transito') || text.includes('lista de espera') || text.includes('admitido')) {
    return 'Aceptado';
  }
  if (text.includes('aceptad')) {
    return 'Aceptado';
  }
  if (text.includes('inscrito')) {
    return 'Inscrito';
  }
  if (text.includes('ganado')) {
    return 'Ganado';
  }
  if (
    text.includes('perdido') ||
    text.includes('abandono') ||
    text.includes('descartado') ||
    text.includes('cancelado') ||
    text.includes('no cumple')
  ) {
    return 'Perdido';
  }
  return phase;
}

function truncate_label(value, max_length = 28) {
  const text = clean(value);
  if (text.length <= max_length) {
    return text;
  }
  return `${text.slice(0, max_length - 3)}...`;
}

function is_cita_or_later(phase) {
  const order = phase_rank(phase);
  const cita_order = phase_rank('Citas');
  if (order === -1 || cita_order === -1) {
    return normalize_text(phase).includes('cita');
  }
  return order >= cita_order;
}

function is_won_student_phase(phase) {
  const normalized = normalize_phase(phase);
  return normalized === 'Ganado' || normalized === 'Inscrito';
}

function resolve_won_phase_strategy(rows) {
  const has_ganado_or_inscrito = rows.some((row) => {
    const normalized = normalize_phase(row.phase);
    return normalized === 'Ganado' || normalized === 'Inscrito';
  });

  if (has_ganado_or_inscrito) {
    return {
      match: (phase) => phase === 'Ganado' || phase === 'Inscrito',
      chart_title: 'Nuevos alumnos ganados por fecha',
      hover_label: 'Ganado/Inscrito',
      empty_message: 'No hay registros en fase Ganado/Inscrito con fecha disponible.'
    };
  }

  return {
    match: (phase) => phase === 'Aceptado',
    chart_title: 'Nuevos alumnos aceptados por fecha',
    hover_label: 'Aceptado',
    empty_message: 'No hay registros en fase Aceptado con fecha disponible.'
  };
}

function render_records(rows) {
  if (!records_grid) {
    return;
  }

  const strategy = resolve_won_phase_strategy(rows);
  const won_rows = rows.filter((row) => strategy.match(normalize_phase(row.phase)));
  const won_dates = won_rows
    .map((row) => row.closed_at || row.updated_at || row.created_at)
    .filter((date) => date instanceof Date && !Number.isNaN(date.getTime()));

  const cards = [
    { period: '1 DÍA', title: 'Récord de inscritos', data: get_record_for_period(won_dates, 'day') },
    { period: '1 SEMANA', title: 'Récord de inscritos', data: get_record_for_period(won_dates, 'week') },
    { period: '1 MES', title: 'Récord de inscritos', data: get_record_for_period(won_dates, 'month') },
    { period: '1 AÑO', title: 'Récord de inscritos', data: get_record_for_period(won_dates, 'year') }
  ];

  records_grid.innerHTML = cards
    .map((card) => `
      <article class="record-card">
        <p class="record-period">${card.period}</p>
        <h3>${card.title}</h3>
        <p class="record-value">${fmt_int(card.data.count)}</p>
        <p class="record-date">Cuándo: ${card.data.label}</p>
      </article>
    `)
    .join('');
}

function get_record_for_period(dates, period) {
  if (!dates.length) {
    return { count: 0, label: 'Sin fecha disponible', rank: -Infinity };
  }

  const map = new Map();
  dates.forEach((date) => {
    const { key, label, rank } = period_key_and_label(date, period);
    const current = map.get(key) || { count: 0, label, rank };
    current.count += 1;
    map.set(key, current);
  });

  let best = { count: 0, label: 'Sin fecha disponible', rank: -Infinity };
  map.forEach((entry) => {
    if (entry.count > best.count || (entry.count === best.count && entry.rank > best.rank)) {
      best = entry;
    }
  });
  return best;
}

function period_key_and_label(date, period) {
  const y = date.getFullYear();
  const m = date.getMonth();
  const d = date.getDate();

  if (period === 'day') {
    const key = `${y}-${m + 1}-${d}`;
    return {
      key,
      label: format_day_with_weekday(date),
      rank: new Date(y, m, d).getTime()
    };
  }

  if (period === 'month') {
    const key = `${y}-${m + 1}`;
    const month_name = date.toLocaleDateString('es-MX', { month: 'long' });
    return {
      key,
      label: `${month_name.charAt(0).toUpperCase()}${month_name.slice(1)} ${y}`,
      rank: new Date(y, m, 1).getTime()
    };
  }

  if (period === 'year') {
    const key = `${y}`;
    return {
      key,
      label: `${y}`,
      rank: new Date(y, 0, 1).getTime()
    };
  }

  const week_start = start_of_week(date);
  const week_end = new Date(week_start);
  week_end.setDate(week_start.getDate() + 6);
  const key = `${week_start.getFullYear()}-${week_start.getMonth() + 1}-${week_start.getDate()}`;
  return {
    key,
    label: `${format_date_only(week_start)} al ${format_date_only(week_end)}`,
    rank: week_start.getTime()
  };
}

function start_of_week(date) {
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const start = new Date(date);
  start.setHours(0, 0, 0, 0);
  start.setDate(date.getDate() + diff);
  return start;
}

function phase_rank(phase) {
  const normalized_phase = normalize_phase(phase);
  return phase_order.findIndex((item) => item === normalized_phase);
}

function parse_number(value) {
  if (value == null) {
    return NaN;
  }
  const text = value.toString().replace(/[^0-9.,-]/g, '').replace(',', '.');
  const num = Number(text);
  return Number.isFinite(num) ? num : NaN;
}

function parse_days(value) {
  const text = normalize_text(value);
  if (!text) {
    return NaN;
  }
  if (text.includes('hoy')) {
    return 0;
  }
  const match = text.match(/(\d+)/);
  return match ? Number(match[1]) : NaN;
}

function parse_date(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function calc_age(date) {
  if (!date) {
    return NaN;
  }
  const now = new Date();
  let age = now.getFullYear() - date.getFullYear();
  const month_diff = now.getMonth() - date.getMonth();
  if (month_diff < 0 || (month_diff === 0 && now.getDate() < date.getDate())) {
    age -= 1;
  }
  return age >= 0 && age <= 30 ? age : NaN;
}

function count_by(rows, fn) {
  return rows.reduce((acc, row) => {
    const key = fn(row) || 'Sin dato';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function top_item(counter) {
  const entries = Object.entries(counter).sort((a, b) => b[1] - a[1]);
  return entries[0] ? { label: entries[0][0], value: entries[0][1] } : { label: 'N/D', value: 0 };
}

function unique_sorted(values, sorter) {
  const unique = [...new Set(values.filter(Boolean))];
  return sorter ? unique.sort(sorter) : unique.sort((a, b) => a.localeCompare(b, 'es'));
}

function phase_sorter(a, b) {
  const ai = phase_order.findIndex((phase) => normalize_text(a).includes(normalize_text(phase)));
  const bi = phase_order.findIndex((phase) => normalize_text(b).includes(normalize_text(phase)));
  if (ai === -1 && bi === -1) {
    return a.localeCompare(b, 'es');
  }
  if (ai === -1) {
    return 1;
  }
  if (bi === -1) {
    return -1;
  }
  return ai - bi;
}

function average(values) {
  const valid = values.filter((value) => Number.isFinite(value));
  if (!valid.length) {
    return NaN;
  }
  return valid.reduce((sum, value) => sum + value, 0) / valid.length;
}

function fmt_int(value) {
  return Number(value || 0).toLocaleString('es-MX');
}

function fmt_number(value, decimals = 0) {
  return Number.isFinite(value) ? value.toFixed(decimals) : 'N/D';
}

function format_date(date) {
  return new Intl.DateTimeFormat('es-MX', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  }).format(date);
}

function format_date_only(date) {
  return new Intl.DateTimeFormat('es-MX', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(date);
}

function format_day_with_weekday(date) {
  const weekday = new Intl.DateTimeFormat('es-MX', { weekday: 'long' }).format(date);
  const weekday_clean = weekday.charAt(0).toUpperCase() + weekday.slice(1);
  return `${weekday_clean}, ${format_date_only(date)}`;
}

function file_stamp(date) {
  const y = date.getFullYear();
  const m = `${date.getMonth() + 1}`.padStart(2, '0');
  const d = `${date.getDate()}`.padStart(2, '0');
  const hh = `${date.getHours()}`.padStart(2, '0');
  const mm = `${date.getMinutes()}`.padStart(2, '0');
  return `${y}${m}${d}_${hh}${mm}`;
}

function file_stamp_full(date) {
  const y = date.getFullYear();
  const m = `${date.getMonth() + 1}`.padStart(2, '0');
  const d = `${date.getDate()}`.padStart(2, '0');
  const hh = `${date.getHours()}`.padStart(2, '0');
  const mm = `${date.getMinutes()}`.padStart(2, '0');
  return `${y}_${m}_${d}_${hh}_${mm}`;
}

function to_local_date_key(date) {
  const y = date.getFullYear();
  const m = `${date.getMonth() + 1}`.padStart(2, '0');
  const d = `${date.getDate()}`.padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function weekday_mon_index(date) {
  const day = date.getDay();
  return day === 0 ? 6 : day - 1;
}

function escape_html(value) {
  return (value || '')
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function limit_map(counter, limit) {
  const entries = Object.entries(counter).sort((a, b) => b[1] - a[1]).slice(0, limit);
  return Object.fromEntries(entries);
}
