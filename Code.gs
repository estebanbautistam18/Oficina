/**
 * NACIONES | OFICINA
 * Backend Google Apps Script independiente. Se conecta por ID a la hoja
 * de Hombría al Máximo para reutilizar sus calendarios y métricas.
 */
const OFFICE = {
  TIMEZONE: 'America/Mexico_City',
  SPREADSHEET_ID: '11VNaHB13tsaJcQS6jJY7GIp8k8sr0zLfxb0YTk7evK0',
  SHEETS: {
    TASKS: 'Oficina - Tareas',
    TOP3: 'Oficina - Top 3',
    RETROS: 'Oficina - Retros 11 Sur',
    INVENTORY: 'Oficina - Insumos',
    TEAM: 'Oficina - Equipos',
    AREAS: 'Oficina - Áreas',
    SETTINGS: 'Oficina - Configuración'
  },
  DEFAULT_AREAS: [
    ['11 Sur', '#B45E4A', 1, true],
    ['Hombría al Máximo', '#8C4545', 2, true],
    ['Eventos / Otras áreas', '#66758A', 3, true]
  ],
  TASK_HEADERS: ['id','titulo','descripcion','area','categoria','completada','fecha_limite','prioridad','responsable','semana','origen','campus','creada_en','actualizada_en'],
  TOP3_HEADERS: ['semana','posicion','objetivo','detalle','completado','actualizado_en'],
  RETRO_HEADERS: ['id','fecha','que_salio_bien','que_salio_mal','que_mejorar','comentarios','correo_enviado','creada_en'],
  INVENTORY_HEADERS: ['id','insumo','cantidad','unidad','estado','ubicacion','notas','actualizado_en'],
  TEAM_HEADERS: ['id','nombre','rol','area','telefono','correo','supervisor_id','activo','notas','actualizado_en'],
  AREA_HEADERS: ['nombre','color','orden','activa'],
  SETTINGS_HEADERS: ['clave','valor']
};

function doGet() {
  ensureOfficeSheets_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Naciones | Oficina')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupOfficePlatform() {
  ensureOfficeSheets_();
  ensureWeeklyHombriaTasks_();
  return { ok: true, message: 'Estructura de Oficina creada y lista.' };
}

function authorizeOfficePlatform() {
  ensureOfficeSheets_();
  // Fuerza la solicitud de permisos para hoja, correo y activadores desde el editor.
  getSpreadsheet_().getName();
  MailApp.getRemainingDailyQuota();
  ScriptApp.getProjectTriggers();
  return { ok: true, message: 'Permisos básicos autorizados.' };
}

function getAppData() {
  ensureOfficeSheets_();
  ensureWeeklyHombriaTasks_();
  const ss = getSpreadsheet_();
  const week = currentWeek_();
  const nextStart = addDays_(week.start, 7);
  const nextEnd = addDays_(nextStart, 6);
  const nextWeek = { key: Utilities.formatDate(nextStart, OFFICE.TIMEZONE, 'yyyy-MM-dd'), start: formatDate_(nextStart, 'yyyy-MM-dd'), end: formatDate_(nextEnd, 'yyyy-MM-dd'), label: formatDate_(nextStart, 'd MMM') + ' – ' + formatDate_(nextEnd, 'd MMM yyyy') };
  return {
    week: { key: week.key, start: formatDate_(week.start, 'yyyy-MM-dd'), end: formatDate_(week.end, 'yyyy-MM-dd'), label: week.label },
    nextWeek: nextWeek,
    tasks: readObjects_(ss.getSheetByName(OFFICE.SHEETS.TASKS), OFFICE.TASK_HEADERS),
    top3: getTop3_(week.key),
    nextTop3: getTop3_(nextWeek.key),
    inventory: readObjects_(ss.getSheetByName(OFFICE.SHEETS.INVENTORY), OFFICE.INVENTORY_HEADERS),
    team: readObjects_(ss.getSheetByName(OFFICE.SHEETS.TEAM), OFFICE.TEAM_HEADERS),
    areas: readObjects_(ss.getSheetByName(OFFICE.SHEETS.AREAS), OFFICE.AREA_HEADERS).filter(x => truthy_(x.activa)),
    retros: readObjects_(ss.getSheetByName(OFFICE.SHEETS.RETROS), OFFICE.RETRO_HEADERS).slice(-8).reverse(),
    settings: getSettings_(),
    hombria: getHombriaData_()
  };
}

/* ------------------------------ TAREAS ------------------------------ */
function createTask(payload) {
  ensureOfficeSheets_();
  const p = payload || {};
  if (!String(p.title || '').trim()) throw new Error('Escribe el nombre de la tarea.');
  const now = new Date();
  const row = [
    Utilities.getUuid(), String(p.title).trim(), String(p.description || '').trim(),
    String(p.area || 'Eventos / Otras áreas').trim(), String(p.category || 'General').trim(),
    false, String(p.dueDate || ''), String(p.priority || 'Media'), String(p.owner || '').trim(),
    String(p.weekKey || ''), String(p.source || 'MANUAL'), String(p.campus || ''), now, now
  ];
  getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TASKS).appendRow(row);
  return { ok: true, id: row[0] };
}

function updateTask(payload) {
  ensureOfficeSheets_();
  const p = payload || {};
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TASKS);
  const data = sheet.getDataRange().getValues();
  const index = data.findIndex((r, i) => i > 0 && String(r[0]) === String(p.id));
  if (index < 1) throw new Error('No encontré esa tarea. Actualiza la pantalla e inténtalo de nuevo.');
  const row = data[index];
  const map = { title: 1, description: 2, area: 3, category: 4, completed: 5, dueDate: 6, priority: 7, owner: 8, campus: 11 };
  Object.keys(map).forEach(k => { if (Object.prototype.hasOwnProperty.call(p, k)) row[map[k]] = p[k]; });
  row[13] = new Date();
  sheet.getRange(index + 1, 1, 1, OFFICE.TASK_HEADERS.length).setValues([row]);
  return { ok: true };
}

function deleteTask(id) {
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TASKS);
  const data = sheet.getDataRange().getValues();
  const index = data.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
  if (index < 1) throw new Error('No encontré esa tarea.');
  sheet.deleteRow(index + 1);
  return { ok: true };
}

function ensureWeeklyHombriaTasks_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(OFFICE.SHEETS.TASKS);
  const week = currentWeek_();
  const rows = sheet.getDataRange().getValues();
  const markerKey = 'hombria_generado_' + week.key;
  const settings = getSettings_();
  if (truthy_(settings[markerKey])) return;
  const exists = rows.some((r, i) => i > 0 && String(r[9]) === week.key && String(r[10]) === 'HOMBRIA_SEMANAL');
  if (exists) {
    getSpreadsheet_().getSheetByName(OFFICE.SHEETS.SETTINGS).appendRow([markerKey, true]);
    return;
  }
  const items = [
    ['Puebla · Alabanza','Invitaciones Puebla','Puebla'],
    ['Puebla · Multimedia','Invitaciones Puebla','Puebla'],
    ['Puebla · Audio','Invitaciones Puebla','Puebla'],
    ['Puebla · Host','Invitaciones Puebla','Puebla'],
    ['Puebla · Orador','Invitaciones Puebla','Puebla'],
    ['Huejotzingo · Alabanza','Invitaciones Huejo','Huejotzingo'],
    ['Huejotzingo · Multimedia','Invitaciones Huejo','Huejotzingo'],
    ['Huejotzingo · Audio','Invitaciones Huejo','Huejotzingo'],
    ['Huejotzingo · Host','Invitaciones Huejo','Huejotzingo'],
    ['Huejotzingo · Orador','Invitaciones Huejo','Huejotzingo'],
    ['Agenda','Organización',''],
    ['Vehículo','Organización','']
  ];
  const now = new Date();
  const due = week.end;
  const values = items.map(x => [Utilities.getUuid(), x[0], '', 'Hombría al Máximo', x[1], false, due, 'Media', '', week.key, 'HOMBRIA_SEMANAL', x[2], now, now]);
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, OFFICE.TASK_HEADERS.length).setValues(values);
  getSpreadsheet_().getSheetByName(OFFICE.SHEETS.SETTINGS).appendRow([markerKey, true]);
}

/* ------------------------------ TOP 3 ------------------------------ */
function getTop3_(weekKey) {
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TOP3);
  let rows = sheet.getDataRange().getValues().slice(1).filter(r => String(r[0]) === weekKey);
  if (!rows.length && weekKey === currentWeek_().key && sheet.getLastRow() === 1) {
    const examples = [
      ['Post Hombría','Retros de coordinadores, calificaciones y status'],
      ['Viaje Chalca','Seguimiento y organización del viaje'],
      ['Pendientes','Noche Naciones, folleto de bienvenida, plataforma Andy y Ricardo, calendario diciembre Esperillas, playeras Somos Naciones y guías Valiosa']
    ];
    const now = new Date();
    sheet.getRange(2, 1, 3, OFFICE.TOP3_HEADERS.length).setValues(examples.map((x,i) => [weekKey, i + 1, x[0], x[1], false, now]));
    rows = sheet.getDataRange().getValues().slice(1).filter(r => String(r[0]) === weekKey);
  }
  const slots = [1,2,3].map(slot => {
    const r = rows.find(x => Number(x[1]) === slot);
    return { slot: slot, title: r ? String(r[2] || '') : '', detail: r ? String(r[3] || '') : '', completed: r ? truthy_(r[4]) : false };
  });
  return slots;
}

function saveTop3(payload) {
  ensureOfficeSheets_();
  const weekKey = String(payload.weekKey || currentWeek_().key);
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (items.length !== 3 || items.some(x => !String(x.title || '').trim())) throw new Error('Completa los tres objetivos del Top 3 antes de guardar.');
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TOP3);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]) === weekKey) sheet.deleteRow(i + 1);
  const now = new Date();
  const values = items.map((x,i) => [weekKey, i + 1, String(x.title || '').trim(), String(x.detail || '').trim(), !!x.completed, now]);
  sheet.getRange(sheet.getLastRow() + 1, 1, 3, OFFICE.TOP3_HEADERS.length).setValues(values);
  return { ok: true };
}

function sendTop3Reminder() {
  ensureOfficeSheets_();
  const settings = getSettings_();
  const email = String(settings.email || Session.getEffectiveUser().getEmail() || '').trim();
  if (!email) throw new Error('Configura tu correo en Naciones | Oficina para recibir el recordatorio.');
  const next = addDays_(currentWeek_().start, 7);
  const nextKey = Utilities.formatDate(next, OFFICE.TIMEZONE, 'yyyy-MM-dd');
  const subject = 'Es momento de definir tu Top 3 · Semana del ' + formatDate_(next, 'dd/MM/yyyy');
  const body = [
    'Hola,', '',
    'Es momento de definir tus tres objetivos principales para la siguiente semana.',
    'Abre Naciones | Oficina y completa tu Top 3.', '',
    'Semana que comienza: ' + formatDate_(next, 'dd/MM/yyyy'),
    'No olvides elegir solo tres prioridades.', '',
    'Naciones | Oficina'
  ].join('\n');
  MailApp.sendEmail(email, subject, body);
  return { ok: true, email: email, weekKey: nextKey };
}

function installTop3Reminder() {
  const settings = getSettings_();
  const email = String(settings.email || Session.getEffectiveUser().getEmail() || '').trim();
  if (!email) throw new Error('Primero guarda tu correo en Configuración.');
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendTop3Reminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendTop3Reminder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(12)
    .nearMinute(0)
    .everyWeeks(1)
    .create();
  return { ok: true, message: 'Recordatorio semanal instalado para los sábados alrededor de las 12:00 p. m.' };
}

/* ------------------------------ RETRO 11 SUR ------------------------------ */
function sendRetroEmail(payload) {
  ensureOfficeSheets_();
  const p = payload || {};
  const settings = getSettings_();
  const email = String(settings.email || Session.getEffectiveUser().getEmail() || '').trim();
  if (!email) throw new Error('Guarda tu correo en Configuración antes de enviar la retroalimentación.');
  const date = p.date ? new Date(p.date + 'T12:00:00') : new Date();
  const monthEs = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  const dateLabel = formatDate_(date, 'dd') + '/' + monthEs[Number(formatDate_(date, 'M')) - 1] + '/' + formatDate_(date, 'yyyy');
  const subject = 'Retro 11 sur - ' + dateLabel;
  const section = (title, value) => title + '\n\n' + linesToBullets_(value);
  const body = [
    section('¿Qué salió bien?', p.good), '',
    section('¿Qué salió mal?', p.bad), '',
    section('¿Qué podemos mejorar?', p.improve),
    p.comments ? '\nComentarios\n\n' + linesToBullets_(p.comments) : ''
  ].join('\n');
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.RETROS);
  sheet.appendRow([Utilities.getUuid(), date, p.good || '', p.bad || '', p.improve || '', p.comments || '', false, new Date()]);
  const retroRow = sheet.getLastRow();
  MailApp.sendEmail({ to: email, subject: subject, body: body, name: 'Naciones | Oficina' });
  sheet.getRange(retroRow, 7).setValue(true);
  return { ok: true, email: email, subject: subject };
}

function linesToBullets_(value) {
  const lines = String(value || '').split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  return lines.length ? lines.map(x => '- ' + x.replace(/^[-•]\s*/, '')).join('\n') : '- Sin comentarios';
}

/* ------------------------------ INSUMOS ------------------------------ */
function saveInventoryItem(payload) {
  ensureOfficeSheets_();
  const p = payload || {};
  if (!String(p.item || '').trim()) throw new Error('Escribe el nombre del insumo.');
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.INVENTORY);
  const values = sheet.getDataRange().getValues();
  const idx = p.id ? values.findIndex((r,i) => i > 0 && String(r[0]) === String(p.id)) : -1;
  const row = [p.id || Utilities.getUuid(), String(p.item).trim(), String(p.quantity || ''), String(p.unit || ''), String(p.status || 'Pendiente'), String(p.location || '11 Sur'), String(p.notes || ''), new Date()];
  if (idx > 0) sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]); else sheet.appendRow(row);
  return { ok: true };
}

function deleteInventoryItem(id) {
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.INVENTORY);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex((r,i) => i > 0 && String(r[0]) === String(id));
  if (idx < 1) throw new Error('No encontré ese insumo.');
  sheet.deleteRow(idx + 1);
  return { ok: true };
}

/* ------------------------------ EQUIPO ------------------------------ */
function saveTeamMember(payload) {
  ensureOfficeSheets_();
  const p = payload || {};
  if (!String(p.name || '').trim()) throw new Error('Escribe el nombre de la persona.');
  if (p.id && String(p.supervisorId || '') === String(p.id)) throw new Error('Una persona no puede ser su propio responsable.');
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TEAM);
  const values = sheet.getDataRange().getValues();
  const idx = p.id ? values.findIndex((r,i) => i > 0 && String(r[0]) === String(p.id)) : -1;
  const row = [p.id || Utilities.getUuid(), String(p.name).trim(), String(p.role || '').trim(), String(p.area || '11 Sur'), String(p.phone || '').trim(), String(p.email || '').trim(), String(p.supervisorId || '').trim(), p.active !== false, String(p.notes || '').trim(), new Date()];
  if (idx > 0) sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]); else sheet.appendRow(row);
  return { ok: true };
}

function deleteTeamMember(id) {
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.TEAM);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex((r,i) => i > 0 && String(r[0]) === String(id));
  if (idx < 1) throw new Error('No encontré a esta persona.');
  sheet.deleteRow(idx + 1);
  return { ok: true };
}

/* ------------------------------ ÁREAS Y AJUSTES ------------------------------ */
function saveArea(payload) {
  ensureOfficeSheets_();
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('Escribe el nombre del área.');
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.AREAS);
  const rows = sheet.getDataRange().getValues();
  if (rows.slice(1).some(r => String(r[0]).toLowerCase() === name.toLowerCase())) throw new Error('Esa área ya existe.');
  sheet.appendRow([name, String(payload.color || '#7B8794'), sheet.getLastRow(), true]);
  return { ok: true };
}

function saveSettings(payload) {
  ensureOfficeSheets_();
  const p = payload || {};
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  Object.keys(p).forEach(key => {
    const idx = values.findIndex((r,i) => i > 0 && String(r[0]) === key);
    if (idx > 0) sheet.getRange(idx + 1, 2).setValue(p[key]); else sheet.appendRow([key, p[key]]);
  });
  if (p.installReminder === true) installTop3Reminder();
  return { ok: true };
}

function getSettings_() {
  const sheet = getSpreadsheet_().getSheetByName(OFFICE.SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  const output = {};
  values.slice(1).forEach(r => output[String(r[0])] = r[1]);
  if (!output.email) output.email = Session.getEffectiveUser().getEmail() || '';
  return output;
}

/* ------------------------------ CALENDARIO Y MÉTRICAS ------------------------------ */
function getHombriaData_() {
  const ss = getSpreadsheet_();
  const calendar = ss.getSheetByName('Calendario Oradores / Host');
  const attendance = ss.getSheetByName('Asistencias');
  const result = { puebla: [], huejo: [], metrics: { puebla: [], huejo: [] }, sourceAvailable: !!calendar };
  if (calendar && calendar.getLastRow() > 1) {
    const lastRow = calendar.getLastRow();
    const left = calendar.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
    const right = calendar.getRange(2, 11, lastRow - 1, 9).getDisplayValues();
    result.puebla = left.filter(r => r.some(v => String(v).trim()) && !/^fecha$/i.test(String(r[0]).trim())).map(r => ({ date:r[0], speaker:r[1], centralIdea:r[2], welcome:r[3], announcements:r[4], prayer:r[5], activity:r[6], owner:r[7], metrics:r[8] }));
    result.huejo = right.filter(r => r.some(v => String(v).trim()) && !/^fecha$/i.test(String(r[0]).trim())).map(r => ({ date:r[0], speaker:r[1], centralIdea:r[2], welcome:r[3], announcements:r[4], prayer:r[5], activity:r[6], owner:r[7], metrics:r[8] }));
  }
  if (attendance) {
    const pDates = attendance.getRange('B11:K11').getDisplayValues()[0];
    const pValues = attendance.getRange('B12:K12').getDisplayValues()[0];
    const hDates = attendance.getRange('B27:K27').getDisplayValues()[0];
    const hValues = attendance.getRange('B28:K28').getDisplayValues()[0];
    result.metrics.puebla = pDates.map((d,i) => ({ date:d, attendance: parseMetric_(pValues[i]) })).filter(x => x.date || x.attendance !== null);
    result.metrics.huejo = hDates.map((d,i) => ({ date:d, attendance: parseMetric_(hValues[i]) })).filter(x => x.date || x.attendance !== null);
  }
  const settings = getSettings_();
  result.metrics.enrolledPuebla = Number(settings.enrolled_puebla || 0);
  result.metrics.enrolledHuejo = Number(settings.enrolled_huejo || 0);
  return result;
}

function parseMetric_(value) {
  const clean = String(value || '').replace(/[^0-9.-]/g, '');
  return clean === '' ? null : Number(clean);
}

/* ------------------------------ UTILIDADES ------------------------------ */
function getSpreadsheet_() {
  return SpreadsheetApp.openById(OFFICE.SPREADSHEET_ID);
}

function ensureOfficeSheets_() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, OFFICE.SHEETS.TASKS, OFFICE.TASK_HEADERS);
  ensureSheet_(ss, OFFICE.SHEETS.TOP3, OFFICE.TOP3_HEADERS);
  ensureSheet_(ss, OFFICE.SHEETS.RETROS, OFFICE.RETRO_HEADERS);
  ensureSheet_(ss, OFFICE.SHEETS.INVENTORY, OFFICE.INVENTORY_HEADERS);
  ensureSheet_(ss, OFFICE.SHEETS.TEAM, OFFICE.TEAM_HEADERS);
  const areas = ensureSheet_(ss, OFFICE.SHEETS.AREAS, OFFICE.AREA_HEADERS);
  ensureSheet_(ss, OFFICE.SHEETS.SETTINGS, OFFICE.SETTINGS_HEADERS);
  if (areas.getLastRow() === 1) {
    areas.getRange(2, 1, OFFICE.DEFAULT_AREAS.length, OFFICE.AREA_HEADERS.length).setValues(OFFICE.DEFAULT_AREAS);
  }
  [OFFICE.SHEETS.TASKS, OFFICE.SHEETS.TOP3, OFFICE.SHEETS.RETROS, OFFICE.SHEETS.INVENTORY, OFFICE.SHEETS.TEAM, OFFICE.SHEETS.AREAS, OFFICE.SHEETS.SETTINGS].forEach(name => {
    const sh = ss.getSheetByName(name);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#F1F3F5');
  });
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function readObjects_(sheet, headers) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.filter(r => r.some(v => v !== '' && v !== null)).map(r => {
    const o = {};
    headers.forEach((h,i) => {
      let value = r[i];
      if (value instanceof Date) value = h === 'fecha_limite' ? Utilities.formatDate(value, OFFICE.TIMEZONE, 'yyyy-MM-dd') : Utilities.formatDate(value, OFFICE.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
      if (h === 'completada' || h === 'completado' || h === 'activa' || h === 'activo' || h === 'correo_enviado') value = truthy_(value);
      o[h] = value;
    });
    return normalizeObject_(o);
  });
}

function normalizeObject_(o) {
  const aliases = {
    id:'id', titulo:'title', descripcion:'description', area:'area', categoria:'category', completada:'completed',
    fecha_limite:'dueDate', prioridad:'priority', responsable:'owner', semana:'weekKey', origen:'source', campus:'campus',
    creada_en:'createdAt', actualizada_en:'updatedAt', insumo:'item', cantidad:'quantity', unidad:'unit', estado:'status',
    ubicacion:'location', notas:'notes', nombre:'name', rol:'role', telefono:'phone', correo:'email', supervisor_id:'supervisorId',
    activo:'active', que_salio_bien:'good', que_salio_mal:'bad', que_mejorar:'improve', comentarios:'comments', correo_enviado:'emailSent'
  };
  const out = {};
  Object.keys(o).forEach(k => out[aliases[k] || k] = o[k]);
  return out;
}

function currentWeek_() {
  const today = new Date();
  const day = Number(Utilities.formatDate(today, OFFICE.TIMEZONE, 'u'));
  const monday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (day - 1), 12, 0, 0);
  const sunday = addDays_(monday, 6);
  return { key: Utilities.formatDate(monday, OFFICE.TIMEZONE, 'yyyy-MM-dd'), start: monday, end: sunday, label: formatDate_(monday, 'd MMM') + ' – ' + formatDate_(sunday, 'd MMM yyyy') };
}

function addDays_(date, count) {
  const d = new Date(date);
  d.setDate(d.getDate() + count);
  return d;
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(new Date(date), OFFICE.TIMEZONE, pattern);
}

function truthy_(value) {
  return value === true || value === 1 || String(value).toLowerCase() === 'true' || String(value).toLowerCase() === 'sí';
}
