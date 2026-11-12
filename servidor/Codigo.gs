/**
 * APP RRHH - Vacaciones · servidor
 * Va dentro de la hoja "APP RRHH - Vacaciones" (Extensiones > Apps Script).
 * La app del móvil le envía peticiones y este código lee y escribe en las pestañas.
 */

const CFG = {
  DIAS_ANUALES: 22,      // días de vacaciones de un año completo
  DIAS_FIJOS: 10,        // días que fija el administrador cada año (no en el año de arranque)
  ANIO_INICIO: 2026,     // año en que se empieza a usar la app
  PIN_PROV: '1234',      // PIN provisional de los empleados nuevos
  HORAS_SESION: 6,       // tiempo que dura una sesión abierta
  DIAS_RECORDAR: 90,     // con "Recordarme en este móvil": días sin volver a pedir el PIN (se renueva al usarla)
  INTENTOS: 5,           // intentos fallidos de PIN antes de bloquear 15 minutos
  DIAS_RECORDATORIO: 15, // aviso al administrador cuando falten estos días para una solicitud pendiente
  HORA_RECORDATORIO: 8   // hora de la revisión diaria
};

// Avisos al móvil (Firebase Cloud Messaging). Identificador del proyecto de Firebase:
const FIREBASE_PROYECTO = 'rrhh-jamones-granadinos';

const HOJA = { emp: 'Empleados', vac: 'Vacaciones', aus: 'Ausencias y bajas', ext: 'Días extra', fest: 'Festivos' };

const ORIGEN = { empleado: 'Pedida por el empleado', admin: 'Días fijos', previa: 'Aprobada antes de la app' };
const ESTADO = { pendiente: 'Pendiente', aprobada: 'Aprobada', rechazada: 'Rechazada' };
const TIPO   = { horas: 'Ausencia por horas', baja: 'Baja', permiso: 'Permiso', falta: 'Falta puntual', retraso: 'Retraso' };
const inv = o => Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k]));
const ORIGEN_I = inv(ORIGEN), ESTADO_I = inv(ESTADO), TIPO_I = inv(TIPO);

/* ================= Entrada ================= */

function doGet() { return salida_({ ok: true, app: 'APP RRHH' }); }

function doPost(e) {
  let r;
  try {
    const p = JSON.parse(e.postData.contents);
    r = despachar_(p);
  } catch (err) {
    r = { ok: false, error: String(err && err.message || err) };
  }
  return salida_(r);
}

function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function despachar_(p) {
  if (p.accion === 'entrar') return entrar_(p);
  if (p.accion === 'salir') { cerrarSesion_(p.token); quitarDispositivo_(p.fcm); return { ok: true }; }
  const ses = sesion_(p.token);
  if (!ses) return { ok: false, error: 'SESION', mensaje: 'La sesión ha caducado. Vuelve a entrar.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let r;
  COLA_AVISOS_ = [];
  try {
    const f = ses.rol === 'admin' ? ADMIN_[p.accion] : EMP_[p.accion];
    if (!f) throw new Error('Acción no permitida.');
    r = f(ses, p) || {};
    if (p.accion !== 'estado') SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  enviarCola_();
  return Object.assign({ ok: true }, r);
}

/* ================= Sesiones y PIN ================= */

function sal_() {
  const pr = PropertiesService.getScriptProperties();
  let s = pr.getProperty('SAL');
  if (!s) { s = Utilities.getUuid(); pr.setProperty('SAL', s); }
  return s;
}
function hash_(pin) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal_() + '|' + pin, Utilities.Charset.UTF_8);
  return 'h:' + Utilities.base64Encode(b);
}
function pinOk_(guardado, pin) {
  guardado = String(guardado || '');
  return guardado.indexOf('h:') === 0 ? guardado === hash_(pin) : guardado === String(pin);
}
function sesion_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const t = cache.get('t:' + token);
  if (t) return JSON.parse(t);
  // sesión recordada en el móvil
  const pr = PropertiesService.getScriptProperties();
  const g = pr.getProperty('s:' + token);
  if (!g) return null;
  const d = JSON.parse(g);
  if (!(d.exp > Date.now())) { pr.deleteProperty('s:' + token); return null; }
  d.exp = Date.now() + CFG.DIAS_RECORDAR * 86400000;
  pr.setProperty('s:' + token, JSON.stringify(d));
  const datos = { rol: d.rol, id: d.id };
  cache.put('t:' + token, JSON.stringify(datos), CFG.HORAS_SESION * 3600);
  return datos;
}
function abrirSesion_(datos, recordar) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('t:' + token, JSON.stringify(datos), CFG.HORAS_SESION * 3600);
  if (recordar) {
    const pr = PropertiesService.getScriptProperties();
    limpiarSesiones_(pr, null);
    pr.setProperty('s:' + token, JSON.stringify(Object.assign({ exp: Date.now() + CFG.DIAS_RECORDAR * 86400000 }, datos)));
  }
  return token;
}
function cerrarSesion_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('t:' + token);
  PropertiesService.getScriptProperties().deleteProperty('s:' + token);
}
/** Borra sesiones recordadas caducadas y, si se indica, todas las de un empleado (PIN restablecido o salida). */
function limpiarSesiones_(pr, idEmpleado) {
  const todo = pr.getProperties(), ahora = Date.now(), cache = CacheService.getScriptCache();
  Object.keys(todo).forEach(k => {
    if (k.indexOf('s:') !== 0) return;
    let d; try { d = JSON.parse(todo[k]); } catch (e) { d = {}; }
    if (!(d.exp > ahora) || (idEmpleado && d.id === idEmpleado)) { pr.deleteProperty(k); cache.remove('t:' + k.slice(2)); }
  });
}
function normDni_(v) { return String(v || '').toUpperCase().replace(/[\s.-]/g, ''); }

function entrar_(p) {
  const cache = CacheService.getScriptCache();
  const clave = 'i:' + (p.modo === 'admin' ? 'admin' : normDni_(p.usuario));
  const fallos = Number(cache.get(clave) || 0);
  if (fallos >= CFG.INTENTOS) return { ok: false, mensaje: 'Demasiados intentos. Espera 15 minutos y vuelve a probar.' };
  const fallo = msg => { cache.put(clave, String(fallos + 1), 900); return { ok: false, mensaje: msg }; };

  if (p.modo === 'admin') {
    const h = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
    if (!h) return { ok: false, mensaje: 'Falta configurar el PIN del administrador. Hazlo desde la hoja: menú App RRHH > Cambiar PIN del administrador.' };
    if (String(p.usuario || '').trim().toLowerCase() !== 'admin' || hash_(p.pin) !== h) return fallo('Usuario o PIN incorrectos.');
    cache.remove(clave);
    return { ok: true, token: abrirSesion_({ rol: 'admin' }, !!p.recordar), rol: 'admin' };
  }

  const dni = normDni_(p.usuario);
  const t = tabla_(HOJA.emp);
  const fila = t.filas.find(f => normDni_(f[2]) === dni);
  if (!fila || !pinOk_(fila[3], p.pin)) return fallo('DNI o PIN incorrectos.');
  if (String(fila[5]).trim() === 'No') return { ok: false, mensaje: 'Este usuario ya no tiene acceso a la app.' };
  cache.remove(clave);
  return { ok: true, token: abrirSesion_({ rol: 'emp', id: fila[0] }, !!p.recordar), rol: 'emp', cambiarPin: String(fila[4]).trim() !== 'Sí' };
}

/* ================= Lectura de la hoja ================= */

function zona_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function hoy_() { return Utilities.formatDate(new Date(), zona_(), 'yyyy-MM-dd'); }
function esFecha_(v) { return v != null && typeof v === 'object' && typeof v.getTime === 'function'; }
function fecha_(v) {
  if (v === '' || v == null) return '';
  if (esFecha_(v)) return Utilities.formatDate(v, zona_(), 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return s;
}
function hora_(v) {
  if (v === '' || v == null) return '';
  if (esFecha_(v)) return Utilities.formatDate(v, zona_(), 'HH:mm');
  return String(v).trim().slice(0, 5);
}
function hoja_(nombre) {
  const h = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!h) throw new Error('No encuentro la pestaña "' + nombre + '" en la hoja.');
  return h;
}
/** Devuelve las filas con datos (sin cabecera) y el número de fila real de cada una. */
function tabla_(nombre) {
  const h = hoja_(nombre);
  const n = h.getLastRow() - 1, c = h.getLastColumn();
  const filas = n > 0 ? h.getRange(2, 1, n, c).getValues() : [];
  const res = [];
  filas.forEach((f, i) => { if (String(f[0]).trim() !== '') { f.fila = i + 2; res.push(f); } });
  return { hoja: h, filas: res };
}
function buscar_(nombre, id) {
  const t = tabla_(nombre);
  const f = t.filas.find(x => String(x[0]) === String(id));
  if (!f) throw new Error('No encuentro ese registro. Puede que alguien lo haya borrado; recarga la app.');
  return { hoja: t.hoja, f };
}
function nuevoId_(letra) { return letra + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function nombre_(id) { const f = tabla_(HOJA.emp).filas.find(x => x[0] === id); return f ? f[1] : ''; }

function leerTodo_() {
  const empleados = [], inicial = {};
  tabla_(HOJA.emp).filas.forEach(f => {
    empleados.push({ id: f[0], nombre: f[1], dni: normDni_(f[2]), pinCambiado: String(f[4]).trim() === 'Sí',
      activo: String(f[5]).trim() !== 'No', salida: fecha_(f[7]) });
    if (f[6] !== '' && f[6] != null) inicial[f[0]] = Number(f[6]);
  });
  const vac = tabla_(HOJA.vac).filas.map(f => {
    const v = { id: String(f[0]), anio: Number(f[1]) || Number(fecha_(f[4]).slice(0, 4)), emp: f[2], desde: fecha_(f[4]), hasta: fecha_(f[5]), medio: !!f[6], dias: Number(f[7]) || 0,
      origen: ORIGEN_I[f[8]] || 'empleado', estado: ESTADO_I[f[9]] || 'pendiente', creada: fecha_(f[10]) };
    if (f[6]) v.franja = String(f[6]).toLowerCase();
    if (f[12]) v.motivoRechazo = f[12];
    return v;
  });
  const aus = tabla_(HOJA.aus).filas.map(f => {
    const a = { id: String(f[0]), emp: f[2], tipo: TIPO_I[f[4]] || 'falta', desde: fecha_(f[5]), hasta: fecha_(f[6]) || null, motivo: f[13] || '' };
    if (a.tipo === 'retraso') { a.prevista = hora_(f[7]); a.llegada = hora_(f[8]); }
    if (a.tipo === 'horas') { a.salida = hora_(f[10]); a.vuelta = hora_(f[11]); a.categoria = f[12] || 'Cita médica';
      a.origen = String(f[14]).trim() === 'Sí' ? 'empleado' : 'admin'; a.visto = String(f[15]).trim() === 'Sí'; }
    return a;
  });
  const extras = tabla_(HOJA.ext).filas.map(f => ({ id: String(f[0]), anio: Number(f[1]), emp: f[2], dias: Number(f[4]) || 0, motivo: f[5] || '' }));
  const festivos = tabla_(HOJA.fest).filas.map(f => ({ fecha: fecha_(f[0]), nombre: f[2], ambito: f[3] || '' })).filter(x => x.fecha);
  return { empleados, inicial, vac, aus, extras, festivos };
}

/* ================= Cálculo de días (igual que en la app) ================= */

function sumar_(s, n) { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function laborables_(desde, hasta, festivos) {
  const fs = new Set(festivos.map(f => f.fecha)); let n = 0;
  for (let s = desde; s <= hasta; s = sumar_(s, 1)) { const w = new Date(s + 'T00:00:00Z').getUTCDay(); if (w > 0 && w < 6 && !fs.has(s)) n++; }
  return n;
}
function saldo_(D, id, a) {
  const fijosObj = a === CFG.ANIO_INICIO ? 0 : CFG.DIAS_FIJOS;
  const vs = D.vac.filter(v => v.emp === id && v.anio === a);   // año al que cuentan los días
  const extra = D.extras.filter(x => x.emp === id && x.anio === a).reduce((t, x) => t + x.dias, 0);
  const base = a === CFG.ANIO_INICIO ? Number(D.inicial[id] || 0) : CFG.DIAS_ANUALES;
  const arrastre = a > CFG.ANIO_INICIO ? sobrante_(D, id, a - 1) : 0;
  const total = base + extra + arrastre;
  const apro = vs.filter(v => v.estado === 'aprobada');
  const fijos = apro.filter(v => v.origen === 'admin').reduce((t, v) => t + v.dias, 0);
  const pedidos = apro.filter(v => v.origen !== 'admin').reduce((t, v) => t + v.dias, 0);
  const pend = vs.filter(v => v.estado === 'pendiente').reduce((t, v) => t + v.dias, 0);
  return { total, fijos, pedidos, pend, pedir: total - Math.max(fijosObj, fijos) - pedidos - pend };
}
function sobrante_(D, id, a) { if (a < CFG.ANIO_INICIO) return 0; const s = saldo_(D, id, a); return Math.max(0, s.total - s.fijos - s.pedidos); }

function calcularDias_(desde, hasta, medio, festivos) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde) || !/^\d{4}-\d{2}-\d{2}$/.test(hasta)) throw new Error('Fechas no válidas.');
  if (hasta < desde) throw new Error('La fecha final es anterior a la inicial.');
  if (desde.slice(0, 4) !== hasta.slice(0, 4)) throw new Error('Las vacaciones tienen que estar dentro del mismo año.');
  if (medio && desde !== hasta) throw new Error('El medio día tiene que ser un solo día.');
  const l = laborables_(desde, hasta, festivos);
  if (!l) throw new Error('En esas fechas no hay ningún día laborable.');
  return medio ? 0.5 : l;
}
function solapa_(D, emp, desde, hasta, franja) {
  return D.vac.find(v => v.emp === emp && v.estado !== 'rechazada' && v.desde <= hasta && v.hasta >= desde &&
    !(franja && v.medio && v.franja && v.franja !== franja));
}

/* ================= Escritura ================= */

function filaVac_(v) {
  return [v.id, v.anio || Number(v.desde.slice(0, 4)), v.emp, nombre_(v.emp), v.desde, v.hasta, v.medio ? (v.franja === 'tarde' ? 'Tarde' : 'Mañana') : '',
    v.dias, ORIGEN[v.origen], ESTADO[v.estado], hoy_(), v.estado === 'pendiente' ? '' : hoy_(), ''];
}
function crearVac_(v) { hoja_(HOJA.vac).appendRow(filaVac_(v)); }

function crearAus_(a) {
  const t = a.tipo;
  hoja_(HOJA.aus).appendRow([a.id, Number(a.desde.slice(0, 4)), a.emp, nombre_(a.emp), TIPO[t], a.desde, a.hasta || '',
    t === 'retraso' ? "'" + a.prevista : '', t === 'retraso' ? "'" + a.llegada : '', t === 'retraso' ? minutos_(a.prevista, a.llegada) : '',
    t === 'horas' ? "'" + a.salida : '', t === 'horas' ? "'" + a.vuelta : '', t === 'horas' ? a.categoria : '', a.motivo || '',
    t === 'horas' ? (a.origen === 'empleado' ? 'Sí' : 'No') : '', t === 'horas' ? (a.visto ? 'Sí' : 'No') : '', hoy_()]);
}
function minutos_(p, l) { const [a, b] = p.split(':').map(Number), [c, d] = l.split(':').map(Number); return (c * 60 + d) - (a * 60 + b); }
function borrarFila_(nombre, id) { const r = buscar_(nombre, id); r.hoja.deleteRow(r.f.fila); return r.f; }
function texto_(s, max) { return String(s || '').trim().slice(0, max || 300); }
function fechaValida_(s) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) throw new Error('Fecha no válida.'); return s; }
function horaValida_(s) { if (!/^\d{2}:\d{2}$/.test(String(s))) throw new Error('Hora no válida.'); return s; }

/* ================= Acciones del empleado ================= */

const EMP_ = {
  estado(ses) {
    const D = leerTodo_(), id = ses.id, yo = D.empleados.find(e => e.id === id);
    if (!yo || !yo.activo) throw new Error('Este usuario ya no tiene acceso a la app.');
    return { hoy: hoy_(), cfg: CFG, festivos: D.festivos,
      empleados: [{ id: yo.id, nombre: yo.nombre, pinCambiado: yo.pinCambiado, activo: true }],
      inicial: D.inicial[id] != null ? { [id]: D.inicial[id] } : {},
      vac: D.vac.filter(v => v.emp === id), extras: D.extras.filter(x => x.emp === id),
      aus: D.aus.filter(a => a.emp === id && a.tipo === 'horas' && a.origen === 'empleado') };
  },
  cambiarPin(ses, p) {
    const pin = String(p.pin || '');
    if (!/^\d{4,8}$/.test(pin)) throw new Error('El PIN tiene que tener entre 4 y 8 cifras, solo números.');
    if (pin === CFG.PIN_PROV) throw new Error('No puedes usar el PIN provisional. Elige otro.');
    const r = buscar_(HOJA.emp, ses.id);
    r.hoja.getRange(r.f.fila, 4, 1, 2).setValues([[hash_(pin), 'Sí']]);
  },
  pedir(ses, p) {
    const D = leerTodo_(), desde = fechaValida_(p.desde), hasta = p.medio ? desde : fechaValida_(p.hasta || p.desde);
    const franja = p.medio ? (p.franja === 'tarde' ? 'tarde' : 'mañana') : null;
    const dias = calcularDias_(desde, hasta, !!p.medio, D.festivos);
    if (desde < hoy_()) throw new Error('No puedes pedir días que ya han pasado.');
    if (solapa_(D, ses.id, desde, hasta, franja)) throw new Error('Ya tienes vacaciones o una solicitud en esas fechas.');
    const s = saldo_(D, ses.id, Number(desde.slice(0, 4)));
    if (dias > s.pedir) throw new Error('No te quedan días suficientes para pedir (' + Math.max(0, s.pedir) + ').');
    crearVac_({ id: nuevoId_('V'), emp: ses.id, desde, hasta, medio: !!p.medio, franja, dias, origen: 'empleado', estado: 'pendiente' });
    avisar_('admin', 'Nueva solicitud de vacaciones', nombreCorto_(ses.id) + ': ' + periodo_(desde, hasta, franja) + ' (' + diasTexto_(dias) + ')');
  },
  anular(ses, p) {
    const r = buscar_(HOJA.vac, p.id);
    if (r.f[2] !== ses.id || r.f[9] !== ESTADO.pendiente) throw new Error('Solo puedes anular tus solicitudes pendientes.');
    r.hoja.deleteRow(r.f.fila);
  },
  aviso(ses, p) {
    const desde = fechaValida_(p.desde), salida = horaValida_(p.salida), vuelta = horaValida_(p.vuelta);
    if (minutos_(salida, vuelta) <= 0) throw new Error('La hora de vuelta tiene que ser posterior a la de salida.');
    if (desde < hoy_()) throw new Error('Ese día ya ha pasado.');
    crearAus_({ id: nuevoId_('A'), emp: ses.id, tipo: 'horas', desde, hasta: desde, salida, vuelta,
      categoria: p.categoria === 'Otro motivo' ? 'Otro motivo' : 'Cita médica', motivo: texto_(p.motivo), origen: 'empleado', visto: false });
    avisar_('admin', 'Aviso de cita', nombreCorto_(ses.id) + ' sale el ' + fechaCorta_(desde) + ' de ' + salida + ' a ' + vuelta);
  },
  registrarAviso(ses, p) { guardarDispositivo_(p.fcm, ses.id); },
  anularAviso(ses, p) {
    const r = buscar_(HOJA.aus, p.id);
    if (r.f[2] !== ses.id || r.f[4] !== TIPO.horas || String(r.f[14]).trim() !== 'Sí' || fecha_(r.f[5]) < hoy_()) throw new Error('No puedes anular ese aviso.');
    r.hoja.deleteRow(r.f.fila);
  }
};

/* ================= Acciones del administrador ================= */

const ADMIN_ = {
  estado() { const D = leerTodo_(); return Object.assign({ hoy: hoy_(), cfg: CFG }, D); },
  aprobar(ses, p) {
    const r = buscar_(HOJA.vac, p.id);
    r.hoja.getRange(r.f.fila, 10, 1, 3).setValues([[ESTADO.aprobada, r.f[10], hoy_()]]);
    avisar_(r.f[2], 'Vacaciones aprobadas', 'Te han aprobado: ' + periodo_(fecha_(r.f[4]), fecha_(r.f[5]), r.f[6] ? String(r.f[6]).toLowerCase() : null));
  },
  rechazar(ses, p) {
    const r = buscar_(HOJA.vac, p.id), motivo = texto_(p.motivo);
    r.hoja.getRange(r.f.fila, 10, 1, 4).setValues([[ESTADO.rechazada, r.f[10], hoy_(), motivo]]);
    avisar_(r.f[2], 'Vacaciones rechazadas', 'No se han aprobado: ' + periodo_(fecha_(r.f[4]), fecha_(r.f[5]), r.f[6] ? String(r.f[6]).toLowerCase() : null) + (motivo ? '. Motivo: ' + motivo : ''));
  },
  registrarAviso(ses, p) { guardarDispositivo_(p.fcm, 'admin'); },
  probarAviso() {
    const n = dispositivos_('admin').length;
    if (!n) throw new Error('Este usuario no tiene ningún móvil con los avisos activados.');
    const res = enviarAhora_('admin', 'Aviso de prueba', 'Si ves esto, los avisos de la app funcionan.');
    if (res.errores.length && !res.enviados) throw new Error('No se ha podido enviar: ' + res.errores[0]);
    return { enviados: res.enviados };
  },
  vacacionesAprobadas(ses, p) {   // días fijos o vacaciones aprobadas antes de la app
    const D = leerTodo_(), desde = fechaValida_(p.desde), hasta = p.medio ? desde : fechaValida_(p.hasta || p.desde);
    const franja = p.medio ? (p.franja === 'tarde' ? 'tarde' : 'mañana') : null;
    const dias = calcularDias_(desde, hasta, !!p.medio, D.festivos);
    const choque = solapa_(D, p.emp, desde, hasta, franja);
    if (choque) throw new Error('Choca con otras vacaciones (' + choque.desde + ' a ' + choque.hasta + ').');
    // Días adelantados: fechas de un año que cuentan como vacaciones del año siguiente
    const anioFechas = Number(desde.slice(0, 4)), anio = Number(p.anio) || anioFechas;
    if (anio !== anioFechas && anio !== anioFechas + 1) throw new Error('Los días solo pueden contar para su año o para el siguiente.');
    crearVac_({ id: nuevoId_('V'), anio, emp: p.emp, desde, hasta, medio: !!p.medio, franja, dias,
      origen: p.origen === 'admin' ? 'admin' : 'previa', estado: 'aprobada' });
  },
  quitarVac(ses, p) { borrarFila_(HOJA.vac, p.id); },
  extra(ses, p) {
    const dias = Number(p.dias); if (!(dias > 0)) throw new Error('Indica los días.');
    hoja_(HOJA.ext).appendRow([nuevoId_('X'), Number(p.anio), p.emp, nombre_(p.emp), dias, texto_(p.motivo), hoy_()]);
  },
  quitarExtra(ses, p) { borrarFila_(HOJA.ext, p.id); },
  inicial(ses, p) {
    const r = buscar_(HOJA.emp, p.emp);
    r.hoja.getRange(r.f.fila, 7).setValue(p.dias === '' || p.dias == null ? '' : Number(p.dias));
  },
  nuevoEmp(ses, p) {
    const dni = normDni_(p.dni), t = tabla_(HOJA.emp);
    if (t.filas.some(f => normDni_(f[2]) === dni)) throw new Error('Ya hay un empleado con ese DNI.');
    const n = t.filas.reduce((m, f) => Math.max(m, Number(String(f[0]).replace(/\D/g, '')) || 0), 0) + 1;
    const id = 'E' + String(n).padStart(2, '0');
    t.hoja.appendRow([id, texto_(p.nombre, 120).toUpperCase(), dni, "'" + CFG.PIN_PROV, 'No', 'Sí', '', '', '']);
    return { id };
  },
  resetPin(ses, p) {
    const r = buscar_(HOJA.emp, p.emp); r.hoja.getRange(r.f.fila, 4, 1, 2).setValues([["'" + CFG.PIN_PROV, 'No']]);
    limpiarSesiones_(PropertiesService.getScriptProperties(), p.emp);   // cierra sus sesiones recordadas
    borrarDispositivosDe_(p.emp);
  },
  salida(ses, p) {
    const fecha = fechaValida_(p.fecha), r = buscar_(HOJA.emp, p.emp);
    limpiarSesiones_(PropertiesService.getScriptProperties(), p.emp);
    borrarDispositivosDe_(p.emp);
    r.hoja.getRange(r.f.fila, 6).setValue('No'); r.hoja.getRange(r.f.fila, 8).setValue(fecha);
    // anula solicitudes pendientes y avisos posteriores a la salida
    const tv = tabla_(HOJA.vac); tv.filas.filter(f => f[2] === p.emp && f[9] === ESTADO.pendiente).map(f => f.fila).sort((a, b) => b - a).forEach(n => tv.hoja.deleteRow(n));
    const ta = tabla_(HOJA.aus); ta.filas.filter(f => f[2] === p.emp && f[4] === TIPO.horas && fecha_(f[5]) > fecha).map(f => f.fila).sort((a, b) => b - a).forEach(n => ta.hoja.deleteRow(n));
  },
  readmitir(ses, p) { const r = buscar_(HOJA.emp, p.emp); r.hoja.getRange(r.f.fila, 6).setValue('Sí'); r.hoja.getRange(r.f.fila, 8).setValue(''); },
  ausencia(ses, p) {
    const t = TIPO[p.tipo] ? p.tipo : null; if (!t) throw new Error('Tipo de registro no válido.');
    const a = { id: nuevoId_('A'), emp: p.emp, tipo: t, desde: fechaValida_(p.desde), hasta: p.hasta ? fechaValida_(p.hasta) : (t === 'baja' ? null : p.desde), motivo: texto_(p.motivo) };
    if (a.hasta && a.hasta < a.desde) throw new Error('La fecha final es anterior a la inicial.');
    if (t === 'retraso') { a.prevista = horaValida_(p.prevista); a.llegada = horaValida_(p.llegada); if (minutos_(a.prevista, a.llegada) <= 0) throw new Error('La hora de llegada tiene que ser posterior a la de entrada.'); }
    if (t === 'horas') { a.salida = horaValida_(p.salida); a.vuelta = horaValida_(p.vuelta); a.hasta = a.desde; a.categoria = p.categoria === 'Otro motivo' ? 'Otro motivo' : 'Cita médica'; a.origen = 'admin'; a.visto = true;
      if (minutos_(a.salida, a.vuelta) <= 0) throw new Error('La hora de vuelta tiene que ser posterior a la de salida.'); }
    crearAus_(a);
  },
  quitarAus(ses, p) { borrarFila_(HOJA.aus, p.id); },
  visto(ses, p) { const r = buscar_(HOJA.aus, p.id); r.hoja.getRange(r.f.fila, 16).setValue('Sí'); },
  cerrarBaja(ses, p) { const r = buscar_(HOJA.aus, p.id); r.hoja.getRange(r.f.fila, 7).setValue(fechaValida_(p.fin)); },
  festivo(ses, p) {
    const fecha = fechaValida_(p.fecha), t = tabla_(HOJA.fest);
    if (t.filas.some(f => fecha_(f[0]) === fecha)) throw new Error('Ese día ya es festivo.');
    const n = t.hoja.getLastRow() + 1;
    t.hoja.getRange(n, 1, 1, 4).setValues([[fecha, '=YEAR(A' + n + ')', texto_(p.nombre, 120), p.ambito || 'Empresa']]);
  },
  quitarFestivo(ses, p) {
    const t = tabla_(HOJA.fest), f = t.filas.find(x => fecha_(x[0]) === p.fecha);
    if (f) t.hoja.deleteRow(f.fila);
  }
};

/* ================= Avisos al móvil ================= */

let COLA_AVISOS_ = [];
const MESES_ = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
function fechaCorta_(s) { return Number(s.slice(8, 10)) + ' ' + MESES_[Number(s.slice(5, 7)) - 1]; }
function periodo_(desde, hasta, franja) {
  if (franja) return 'medio día (' + franja + ') el ' + fechaCorta_(desde);
  return desde === hasta ? 'el ' + fechaCorta_(desde) : 'del ' + fechaCorta_(desde) + ' al ' + fechaCorta_(hasta);
}
function diasTexto_(n) { return String(n).replace('.', ',') + (n === 1 ? ' día' : ' días'); }
function nombreCorto_(id) {
  const n = String(nombre_(id) || id).toLowerCase().replace(/(^|\s)(\S)/g, (m, a, l) => a + l.toUpperCase());
  return n.split(' ').slice(0, 2).join(' ');
}

/** Dispositivos guardados: propiedad 'd:<huella>' = {fcm, quien, fecha}. */
function claveDispositivo_(fcm) {
  return 'd:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(fcm)));
}
function guardarDispositivo_(fcm, quien) {
  if (!fcm || String(fcm).length < 20) throw new Error('No se ha podido activar los avisos en este móvil.');
  PropertiesService.getScriptProperties().setProperty(claveDispositivo_(fcm), JSON.stringify({ fcm: String(fcm), quien, fecha: hoy_() }));
}
function quitarDispositivo_(fcm) { if (fcm) PropertiesService.getScriptProperties().deleteProperty(claveDispositivo_(fcm)); }
function dispositivos_(quien) {
  const todo = PropertiesService.getScriptProperties().getProperties(), r = [];
  Object.keys(todo).forEach(k => { if (k.indexOf('d:') !== 0) return; try { const d = JSON.parse(todo[k]); if (d.quien === quien) r.push(d.fcm); } catch (e) {} });
  return r;
}
function borrarDispositivosDe_(quien) {
  const pr = PropertiesService.getScriptProperties(), todo = pr.getProperties();
  Object.keys(todo).forEach(k => { if (k.indexOf('d:') !== 0) return; try { if (JSON.parse(todo[k]).quien === quien) pr.deleteProperty(k); } catch (e) {} });
}

/** Deja el aviso preparado; se envía al acabar la acción. */
function avisar_(quien, titulo, cuerpo) { COLA_AVISOS_.push([quien, titulo, cuerpo]); }
function enviarCola_() {
  const cola = COLA_AVISOS_; COLA_AVISOS_ = [];
  cola.forEach(a => { try { enviarAhora_(a[0], a[1], a[2]); } catch (e) { console.warn('Aviso no enviado: ' + e); } });
}
function enviarAhora_(quien, titulo, cuerpo) {
  const res = { enviados: 0, errores: [] };
  if (!FIREBASE_PROYECTO || FIREBASE_PROYECTO.indexOf('PEGA_AQUI') === 0) { res.errores.push('falta el identificador del proyecto de Firebase en el código'); return res; }
  const tokens = dispositivos_(quien);
  if (!tokens.length) return res;
  const url = 'https://fcm.googleapis.com/v1/projects/' + FIREBASE_PROYECTO + '/messages:send';
  const cab = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'x-goog-user-project': FIREBASE_PROYECTO };
  const peticiones = tokens.map(t => ({
    url, method: 'post', contentType: 'application/json', headers: cab, muteHttpExceptions: true,
    payload: JSON.stringify({ message: { token: t, data: { titulo, cuerpo }, webpush: { headers: { Urgency: 'high', TTL: '86400' } } } })
  }));
  UrlFetchApp.fetchAll(peticiones).forEach((resp, i) => {
    const c = resp.getResponseCode();
    if (c === 200) { res.enviados++; return; }
    const txt = resp.getContentText();
    if (c === 404 || /UNREGISTERED|INVALID_ARGUMENT/.test(txt)) { quitarDispositivo_(tokens[i]); }   // móvil que ya no existe
    let msg = txt; try { msg = JSON.parse(txt).error.message; } catch (e) {}
    res.errores.push(c + ' ' + msg);
  });
  return res;
}

/* ================= Recordatorio diario ================= */

/**
 * Se ejecuta sola cada mañana. Avisa al administrador de las solicitudes que siguen pendientes
 * cuando faltan CFG.DIAS_RECORDATORIO días (o menos) para que empiecen. Cada solicitud se recuerda una sola vez.
 */
function recordatorioDiario() {
  const pr = PropertiesService.getScriptProperties(), hoy = hoy_(), limite = sumar_(hoy, CFG.DIAS_RECORDATORIO);
  const pendientes = leerTodo_().vac.filter(v => v.estado === 'pendiente' && v.desde >= hoy && v.desde <= limite);
  pendientes.forEach(v => {
    const clave = 'r:' + v.id;
    if (pr.getProperty(clave)) return;
    const faltan = Math.round((new Date(v.desde + 'T00:00:00Z') - new Date(hoy + 'T00:00:00Z')) / 86400000);
    const cuando = faltan === 0 ? 'empiezan hoy' : faltan === 1 ? 'empiezan mañana' : 'faltan ' + faltan + ' días';
    const res = enviarAhora_('admin', 'Vacaciones sin decidir: ' + cuando,
      nombreCorto_(v.emp) + ': ' + periodo_(v.desde, v.hasta, v.franja || null) + ' (' + diasTexto_(v.dias) + '). Apruébalas o recházalas en la app.');
    if (res.enviados) pr.setProperty(clave, hoy);
  });
  // limpia marcas de solicitudes antiguas
  Object.keys(pr.getProperties()).forEach(k => { if (k.indexOf('r:') === 0 && pr.getProperty(k) < sumar_(hoy, -60)) pr.deleteProperty(k); });
}

/** Deja programada la revisión diaria (una sola vez; si ya existía, la reemplaza). */
function instalarRecordatorio() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'recordatorioDiario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('recordatorioDiario').timeBased().everyDays(1).atHour(CFG.HORA_RECORDATORIO).inTimezone('Europe/Madrid').create();
  const msg = 'Recordatorio activado: cada día hacia las ' + CFG.HORA_RECORDATORIO + ':00 se revisan las solicitudes pendientes.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

/** Ejecútala una vez desde el editor (Ejecutar) para dar todos los permisos, también el de enviar avisos. */
function autorizar() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  UrlFetchApp.fetch('https://www.google.com');
  console.log('Permisos correctos.');
}

/* ================= Menú en la hoja ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('App RRHH')
    .addItem('Cambiar PIN del administrador', 'cambiarPinAdmin')
    .addItem('Activar recordatorio diario de solicitudes', 'instalarRecordatorio')
    .addToUi();
}

function cambiarPinAdmin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('PIN del administrador', 'Escribe el PIN con el que entrarás como administrador (de 4 a 8 cifras). El usuario es "admin".', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pin = r.getResponseText().trim();
  if (!/^\d{4,8}$/.test(pin)) { ui.alert('El PIN tiene que tener entre 4 y 8 cifras, solo números. No se ha cambiado.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', hash_(pin));
  ui.alert('PIN del administrador guardado. Entra en la app con el usuario "admin" y este PIN.');
}
