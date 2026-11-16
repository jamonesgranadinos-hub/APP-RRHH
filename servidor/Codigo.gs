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

// Versión de este código. La hoja se actualiza sola cuando en GitHub hay una versión distinta.
const VERSION_SERVIDOR = '2026-11-16 b';
const REPO_RAW = 'https://raw.githubusercontent.com/jamonesgranadinos-hub/app-rrhh/main/servidor/';

const inv = o => Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k]));
const HOJA = { emp: 'Empleados', vac: 'Vacaciones', aus: 'Ausencias y bajas', ext: 'Días extra', fest: 'Festivos', reg: 'Registro horario', snom: 'Peticiones de nóminas' };

// Nóminas: carpeta RRHH › Nóminas del Drive (con subcarpetas por año y por mes).
// Cada PDF se llama "2026-10 Nómina - Apellidos Nombre.pdf" y se empareja con el empleado que tenga exactamente esos nombres.
const NOMINAS = {
  CARPETA: '1Se4AOpOY-szDwxvQY1dZWejNgTR4uUg9',
  MESES: 12,        // el empleado ve las nóminas de los últimos 12 meses
  DIAS_ACCESO: 7,   // cuando pide las antiguas y el administrador le da acceso, las ve durante estos días
  HORAS_CACHE: 2    // la lista de PDFs se guarda este tiempo (cada hora se renueva sola y se puede forzar desde la app)
};
const ESTADO_SNOM = { pendiente: 'Pendiente', acceso: 'Acceso dado', atendida: 'Atendida' };
const ESTADO_SNOM_I = inv(ESTADO_SNOM);
const ORIGEN_REG = { movil: 'Fichado en el móvil', admin: 'Apuntado por el administrador', corregido: 'Corregido por el administrador' };
const ORIGEN_REG_I = Object.fromEntries(Object.entries(ORIGEN_REG).map(([k, v]) => [v, k]));

const ORIGEN = { empleado: 'Pedida por el empleado', admin: 'Días fijos', previa: 'Aprobada antes de la app', dada: 'Dada por el administrador', regalo: 'Día regalado' };
const ESTADO = { pendiente: 'Pendiente', aprobada: 'Aprobada', rechazada: 'Rechazada' };
const TIPO   = { horas: 'Ausencia por horas', baja: 'Baja', permiso: 'Permiso', falta: 'Falta puntual', retraso: 'Retraso' };
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
  const f = ses.rol === 'admin' ? ADMIN_[p.accion] : EMP_[p.accion];
  if (!f) throw new Error('Acción no permitida.');
  // leer no necesita esperar a nadie; solo las escrituras van de una en una
  const soloLectura = ['estado', 'nominas', 'verNomina'].indexOf(p.accion) >= 0;
  const lock = soloLectura ? null : LockService.getScriptLock();
  if (lock) lock.waitLock(20000);
  let r;
  COLA_AVISOS_ = [];
  try {
    r = f(ses, p) || {};
    if (!soloLectura) SpreadsheetApp.flush();
  } finally {
    if (lock) lock.releaseLock();
  }
  // devuelve ya los datos actualizados para que la app no tenga que pedirlos otra vez
  if (p.conEstado && !soloLectura) r.datos = Object.assign({ ok: true }, (ses.rol === 'admin' ? ADMIN_ : EMP_).estado(ses, {}));
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
let CONTADOR_ID_ = 0;
function nuevoId_(letra) { return letra + Date.now().toString(36).toUpperCase() + (CONTADOR_ID_++).toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function nombre_(id) { const f = tabla_(HOJA.emp).filas.find(x => x[0] === id); return f ? f[1] : ''; }

function leerTodo_() {
  const empleados = [], inicial = {};
  tabla_(HOJA.emp).filas.forEach(f => {
    empleados.push({ id: f[0], nombre: f[1], dni: normDni_(f[2]), pinCambiado: String(f[4]).trim() === 'Sí',
      activo: String(f[5]).trim() !== 'No', salida: fecha_(f[7]), horario: String(f[9] || '').trim() === 'Sí' });
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
  const apro = vs.filter(v => v.estado === 'aprobada' && v.origen !== 'regalo');   // los días regalados no se descuentan
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
    return { hoy: hoy_(), cfg: CFG, festivos: D.festivos, modulos: { nominas: nominasListas_() },
      empleados: [{ id: yo.id, nombre: yo.nombre, pinCambiado: yo.pinCambiado, activo: true, horario: yo.horario }],
      registro: yo.horario ? leerRegistro_(sumar_(hoy_(), -70)).filter(r => r.emp === id) : [],
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
  fichar(ses, p) {
    const yo = tabla_(HOJA.emp).filas.find(f => f[0] === ses.id);
    if (!yo || String(yo[9] || '').trim() !== 'Sí') throw new Error('No tienes acceso al registro horario.');
    const ahora = new Date(), fecha = hoy_(), hora = Utilities.formatDate(ahora, zona_(), 'HH:mm');
    const t = tabla_(hojaRegistro_().getName());
    const abiertos = t.filas.filter(f => f[2] === ses.id && !hora_(f[6])).sort((a, b) => (fecha_(b[4]) + hora_(b[5])).localeCompare(fecha_(a[4]) + hora_(a[5])));
    const abiertoHoy = abiertos.find(f => fecha_(f[4]) === fecha);
    if (p.tipo === 'salida') {
      if (!abiertoHoy) throw new Error('No tienes ninguna entrada abierta hoy. Ficha primero la entrada.');
      if (minutos_(hora_(abiertoHoy[5]), hora) < 1) throw new Error('Acabas de fichar la entrada. Espera un minuto para fichar la salida.');
      t.hoja.getRange(abiertoHoy.fila, 7, 1, 2).setValues([["'" + hora, horas_(hora_(abiertoHoy[5]), hora)]]);
      return { hora };
    }
    if (abiertoHoy) throw new Error('Ya tienes una entrada abierta desde las ' + hora_(abiertoHoy[5]) + '. Ficha la salida.');
    t.hoja.appendRow([nuevoId_('R'), Number(fecha.slice(0, 4)), ses.id, nombre_(ses.id), fecha, "'" + hora, '', '', ORIGEN_REG.movil, '', fecha, '']);
    return { hora };
  },
  /** Solo las nóminas de quien ha entrado: nunca se manda nada de otro empleado. */
  nominas(ses) {
    comprobarNominas_();
    const mias = nominasDe_(ses.id), lim = limiteNominas_(), acceso = accesoNominas_(ses.id);
    return { limite: lim, acceso: acceso || null,
      lista: mias.filter(x => acceso || x.mes >= lim).map(x => ({ id: x.id, mes: x.mes, extra: x.extra })),
      anteriores: mias.filter(x => !(x.mes >= lim)).length,
      peticiones: peticionesNominas_().filter(x => x.emp === ses.id).slice(-10).reverse().map(x => ({ id: x.id, texto: x.texto, estado: x.estado, fecha: x.fecha, hasta: x.hasta, nota: x.nota })) };
  },
  verNomina(ses, p) {
    comprobarNominas_();
    const x = nominasDe_(ses.id).find(n => n.id === String(p.id || ''));
    if (!x) throw new Error('No puedes ver esa nómina.');
    if (!(x.mes >= limiteNominas_()) && !accesoNominas_(ses.id)) throw new Error('Esa nómina tiene más de ' + NOMINAS.MESES + ' meses. Pídela al administrador.');
    return pdfNomina_(x);
  },
  pedirNominas(ses, p) {
    const texto = texto_(p.texto, 300);
    if (!texto) throw new Error('Escribe qué nóminas necesitas.');
    if (peticionesNominas_().filter(x => x.emp === ses.id && x.estado === 'pendiente').length >= 3) throw new Error('Ya tienes peticiones sin contestar. Espera a que el administrador las atienda.');
    hojaPeticionesNominas_().appendRow([nuevoId_('N'), Number(hoy_().slice(0, 4)), ses.id, nombre_(ses.id), texto, ESTADO_SNOM.pendiente, hoy_(), '', '', '']);
    avisar_('admin', 'Petición de nóminas', nombreCorto_(ses.id) + ': ' + texto);
  },
  anularAviso(ses, p) {
    const r = buscar_(HOJA.aus, p.id);
    if (r.f[2] !== ses.id || r.f[4] !== TIPO.horas || String(r.f[14]).trim() !== 'Sí' || fecha_(r.f[5]) < hoy_()) throw new Error('No puedes anular ese aviso.');
    r.hoja.deleteRow(r.f.fila);
  }
};

/* ================= Acciones del administrador ================= */

const ADMIN_ = {
  estado() {
    const D = leerTodo_();
    return Object.assign({ hoy: hoy_(), cfg: CFG, registro: leerRegistro_(sumar_(hoy_(), -400)),
      modulos: { nominas: nominasListas_() }, nomPend: peticionesNominas_().filter(x => x.estado === 'pendiente').length }, D);
  },
  registroAcceso(ses, p) { const r = buscar_(HOJA.emp, p.emp); r.hoja.getRange(r.f.fila, 10).setValue(p.si ? 'Sí' : 'No'); },
  registroGuardar(ses, p) {
    const fecha = fechaValida_(p.fecha), entrada = horaValida_(p.entrada), salida = p.salida ? horaValida_(p.salida) : '';
    const nota = texto_(p.nota);
    if (p.id) {
      const r = buscar_(HOJA.reg, p.id), antes = [hora_(r.f[5]), hora_(r.f[6])];
      const cambio = 'Antes: ' + antes[0] + '–' + (antes[1] || 'sin salida');
      const notaFinal = [String(r.f[9] || ''), cambio, nota].filter(Boolean).join(' · ');
      r.hoja.getRange(r.f.fila, 5, 1, 8).setValues([[fecha, "'" + entrada, salida ? "'" + salida : '', salida ? horas_(entrada, salida) : '',
        ORIGEN_REG.corregido, notaFinal, r.f[10], hoy_()]]);
    } else {
      hojaRegistro_().appendRow([nuevoId_('R'), Number(fecha.slice(0, 4)), p.emp, nombre_(p.emp), fecha, "'" + entrada, salida ? "'" + salida : '',
        salida ? horas_(entrada, salida) : '', ORIGEN_REG.admin, nota, hoy_(), '']);
    }
  },
  registroBorrar(ses, p) { borrarFila_(HOJA.reg, p.id); },
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
  /** Da el mismo periodo de vacaciones a varios empleados a la vez (puentes, cierres...). */
  vacacionesVarios(ses, p) {
    const D = leerTodo_(), desde = fechaValida_(p.desde), hasta = fechaValida_(p.hasta || p.desde);
    const dias = calcularDias_(desde, hasta, false, D.festivos);
    const origen = ['admin', 'dada', 'regalo'].indexOf(p.tipo) >= 0 ? p.tipo : null;
    if (!origen) throw new Error('Elige si se descuenta o es un día regalado.');
    const emps = (p.emps || []).filter(id => D.empleados.some(e => e.id === id && e.activo));
    if (!emps.length) throw new Error('Elige al menos un empleado.');
    const filas = [], saltados = [];
    emps.forEach(id => {
      if (solapa_(D, id, desde, hasta, null)) { saltados.push(nombreCorto_(id)); return; }
      filas.push(filaVac_({ id: nuevoId_('V'), emp: id, desde, hasta, medio: false, dias, origen, estado: 'aprobada' }));
      avisar_(id, origen === 'regalo' ? 'Día libre' : 'Te han dado vacaciones',
        (origen === 'regalo' ? 'Libras ' : 'Tienes vacaciones ') + periodo_(desde, hasta, null) + (origen === 'regalo' ? '. No se descuenta de tus días.' : '.'));
    });
    if (filas.length) { const h = hoja_(HOJA.vac); h.getRange(h.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas); }
    return { creadas: filas.length, saltados };
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
    t.hoja.appendRow([id, texto_(p.nombre, 120).toUpperCase(), dni, "'" + CFG.PIN_PROV, 'No', 'Sí', '', '', '', 'No']);
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
  },
  nominas(ses, p) {
    comprobarNominas_();
    const lista = nominasAsignadas_(indiceNominas_(!!p.forzar), empleadosNominas_());
    return { limite: limiteNominas_(), actualizado: CacheService.getScriptCache().get('nomhora') || '',
      lista: lista.map(x => ({ id: x.id, mes: x.mes, extra: x.extra, nombre: x.nombre, emp: x.emp, dudoso: x.dudoso })),
      apartados: nombresApartados_(),
      alias: empleadosNominas_().filter(e => e.alias.length).map(e => ({ emp: e.id, alias: e.alias })),
      peticiones: peticionesNominas_().slice(-60).reverse() };
  },
  verNomina(ses, p) {
    comprobarNominas_();
    const x = indiceNominas_(false).find(n => n.id === String(p.id || ''));
    if (!x) throw new Error('No encuentro esa nómina. Pulsa "Buscar nóminas nuevas".');
    return pdfNomina_(x);
  },
  /** Para nóminas que no casan solas (nombre escrito distinto): se apunta en la ficha "Nombre en las nóminas". */
  nominaAsignar(ses, p) {
    const nombre = texto_(p.nombre, 120), r = buscar_(HOJA.emp, p.emp);
    if (!claveNombre_(nombre)) throw new Error('Falta el nombre de la nómina.');
    const alias = String(r.f[10] || '').split('|').map(x => x.trim()).filter(Boolean);
    const quedan = p.quitar ? alias.filter(a => claveNombre_(a) !== claveNombre_(nombre)) : alias.concat(alias.some(a => claveNombre_(a) === claveNombre_(nombre)) ? [] : [nombre]);
    if (!String(r.hoja.getRange(1, 11).getValue() || '').trim()) r.hoja.getRange(1, 11).setValue('Nombre en las nóminas');
    r.hoja.getRange(r.f.fila, 11).setValue(quedan.join(' | '));
  },
  /** Nombres que no entran en la app (ex-empleados, jubilaciones...): sus nóminas no las ve nadie y no se avisa de ellas. */
  nominaApartar(ses, p) {
    const nombre = texto_(p.nombre, 120), k = claveNombre_(nombre);
    if (!k) throw new Error('Falta el nombre.');
    const lista = nombresApartados_().filter(n => claveNombre_(n) !== k);
    if (!p.quitar) lista.push(nombre);
    PropertiesService.getScriptProperties().setProperty('NOM_APARTADOS', JSON.stringify(lista));
  },
  nominaAcceso(ses, p) {
    const r = buscar_(HOJA.snom, p.id), hasta = sumar_(hoy_(), NOMINAS.DIAS_ACCESO);
    r.hoja.getRange(r.f.fila, 6, 1, 5).setValues([[ESTADO_SNOM.acceso, r.f[6], hoy_(), hasta, texto_(p.nota)]]);
    avisar_(r.f[2], 'Nóminas anteriores', 'Ya puedes ver todas tus nóminas en la app hasta el ' + fechaCorta_(hasta) + '.');
  },
  nominaAtendida(ses, p) {
    const r = buscar_(HOJA.snom, p.id), nota = texto_(p.nota);
    r.hoja.getRange(r.f.fila, 6, 1, 5).setValues([[ESTADO_SNOM.atendida, r.f[6], hoy_(), '', nota]]);
    avisar_(r.f[2], 'Tu petición de nóminas', 'El administrador la ha atendido' + (nota ? ': ' + nota : '.'));
  }
};

/* ================= Nóminas ================= */

function carpetaNominas_() { return PropertiesService.getScriptProperties().getProperty('CARPETA_NOMINAS') || NOMINAS.CARPETA; }
/** Se marca al ejecutar "autorizar" con éxito: hasta entonces los empleados no ven el botón. */
function nominasListas_() { return PropertiesService.getScriptProperties().getProperty('NOMINAS_LISTO') === '1'; }
function comprobarNominas_() {
  if (!nominasListas_()) throw new Error('Las nóminas aún no están activadas: en Apps Script elige la función "autorizar" y pulsa Ejecutar.');
}
/** "Gutiérrez Heras, Marcial" y "MARCIAL GUTIERREZ HERAS" dan lo mismo: sin tildes, mayúsculas y palabras ordenadas. */
function claveNombre_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean).sort().join(' ');
}
/** Saca mes, si es extra y el nombre de la persona del nombre del archivo (y de sus carpetas). */
function analizarNomina_(archivo, ruta) {
  const base = String(archivo).replace(/\.pdf$/i, '').trim();
  const i = base.lastIndexOf(' - ');
  const persona = i >= 0 ? base.slice(i + 3).trim() : '';
  const cabeza = i >= 0 ? base.slice(0, i) : base;
  const todo = [cabeza].concat((ruta || []).slice().reverse());
  let mes = '';
  for (const t of todo) { const m = String(t).match(/(20\d{2})[-_ .](0[1-9]|1[0-2])/); if (m) { mes = m[1] + '-' + m[2]; break; } }
  const extra = /extra/i.test(cabeza) || (ruta || []).some(t => /extra/i.test(t));
  return { mes, extra, nombre: persona, clave: claveNombre_(persona) };
}
/** Recorre la carpeta de nóminas (sin abrir los PDF) y guarda la lista un rato. */
function indiceNominas_(forzar) {
  if (!forzar) { const g = cacheLeer_('nomidx'); if (g) return JSON.parse(g); }
  const out = [];
  const visitar = (carpeta, ruta, prof) => {
    const fs = carpeta.searchFiles('trashed = false and mimeType = "application/pdf"');
    while (fs.hasNext()) { const f = fs.next(); out.push(Object.assign({ id: f.getId() }, analizarNomina_(f.getName(), ruta))); }
    if (prof >= 4) return;
    const cs = carpeta.searchFolders('trashed = false');
    while (cs.hasNext()) { const c = cs.next(), n = c.getName(); if (/whatsapp/i.test(n)) continue; visitar(c, ruta.concat(n), prof + 1); }
  };
  visitar(DriveApp.getFolderById(carpetaNominas_()), [], 0);
  out.sort((a, b) => (b.mes + (b.extra ? 'x' : '')).localeCompare(a.mes + (a.extra ? 'x' : '')) || a.nombre.localeCompare(b.nombre));
  cacheGuardar_('nomidx', JSON.stringify(out), NOMINAS.HORAS_CACHE * 3600);
  CacheService.getScriptCache().put('nomhora', Utilities.formatDate(new Date(), zona_(), 'yyyy-MM-dd HH:mm'), NOMINAS.HORAS_CACHE * 3600);
  return out;
}
function empleadosNominas_() {
  return tabla_(HOJA.emp).filas.map(f => ({ id: f[0], nombre: f[1], alias: String(f[10] || '').split('|').map(x => x.trim()).filter(Boolean) }));
}
/** Une cada nómina con su empleado. Si el nombre no casa exacto con UNO solo, la nómina no es de nadie (nadie la ve). */
function nominasAsignadas_(indice, emps) {
  const mapa = {};
  emps.forEach(e => [e.nombre].concat(e.alias).forEach(n => {
    const k = claveNombre_(n); if (!k) return;
    mapa[k] = mapa[k] || []; if (mapa[k].indexOf(e.id) < 0) mapa[k].push(e.id);
  }));
  return indice.map(x => { const ids = (x.clave && mapa[x.clave]) || []; return Object.assign({}, x, { emp: ids.length === 1 ? ids[0] : null, dudoso: ids.length > 1 }); });
}
function nombresApartados_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('NOM_APARTADOS') || '[]'); } catch (e) { return []; } }
function nominasDe_(id) { return nominasAsignadas_(indiceNominas_(false), empleadosNominas_()).filter(x => x.emp === id); }
function limiteNominas_() {
  const h = hoy_(); let y = Number(h.slice(0, 4)), m = Number(h.slice(5, 7)) - NOMINAS.MESES;
  while (m < 1) { m += 12; y--; }
  return y + '-' + String(m).padStart(2, '0');
}
/** Fecha hasta la que el empleado puede ver también las nóminas antiguas (o '' si no). */
function accesoNominas_(id) {
  const hoy = hoy_();
  return peticionesNominas_().filter(x => x.emp === id && x.estado === 'acceso' && x.hasta >= hoy).map(x => x.hasta).sort().pop() || '';
}
/** Abre el PDF (solo después de comprobar de quién es) y lo manda dentro de la respuesta: sin enlaces públicos. */
function pdfNomina_(x) {
  const f = DriveApp.getFileById(x.id);
  if (f.isTrashed() || analizarNomina_(f.getName(), []).clave !== x.clave) {
    cacheBorrar_('nomidx');
    throw new Error('Esa nómina ha cambiado en el Drive. Vuelve a entrar en Nóminas.');
  }
  const bytes = f.getBlob().getBytes();
  if (bytes.length > 15 * 1024 * 1024) throw new Error('El PDF es demasiado grande para verlo en la app.');
  return { mes: x.mes, extra: x.extra, pdf: Utilities.base64Encode(bytes) };
}
function hojaPeticionesNominas_() {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(HOJA.snom);
  if (!h) {
    h = ss.insertSheet(HOJA.snom);
    h.getRange(1, 1, 1, 10).setValues([['ID', 'Año', 'ID empleado', 'Empleado', 'Qué necesita', 'Estado', 'Pedida el', 'Resuelta el', 'Puede verlas hasta', 'Nota del administrador']]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}
function peticionesNominas_() {
  if (!SpreadsheetApp.getActive().getSheetByName(HOJA.snom)) return [];
  return tabla_(HOJA.snom).filas.map(f => ({ id: String(f[0]), emp: f[2], texto: f[4] || '', estado: ESTADO_SNOM_I[f[5]] || 'pendiente',
    fecha: fecha_(f[6]), resuelta: fecha_(f[7]), hasta: fecha_(f[8]), nota: f[9] || '' }));
}
/** Cada hora: avisa a cada empleado cuando aparece una nómina suya nueva. */
function avisarNominasNuevas_() {
  if (!nominasListas_()) return;
  const pr = PropertiesService.getScriptProperties(), previo = pr.getProperty('NOM_ULTIMAS');
  const ult = JSON.parse(previo || '{}'), lim = limiteNominas_(), ahora = {};
  nominasAsignadas_(indiceNominas_(true), empleadosNominas_()).forEach(x => {
    if (!x.emp || !x.mes) return;
    const k = x.mes + (x.extra ? 'x' : '');
    if (!ahora[x.emp] || k > ahora[x.emp]) ahora[x.emp] = k;
  });
  Object.keys(ahora).forEach(e => {
    const k = ahora[e];
    if (previo && (!ult[e] || k > ult[e]) && k.slice(0, 7) >= lim) {
      try { enviarAhora_(e, 'Nómina nueva', 'Ya tienes en la app tu nómina ' + (k.length > 7 ? 'extra de ' : 'de ') + nombreMes_(k.slice(0, 7)) + '.'); } catch (err) { console.warn(err); }
    }
    ult[e] = k;
  });
  pr.setProperty('NOM_ULTIMAS', JSON.stringify(ult));
}
const MESES_LARGOS_ = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
function nombreMes_(m) { return MESES_LARGOS_[Number(m.slice(5, 7)) - 1] + ' ' + m.slice(0, 4); }

/** La caché admite textos de unos 100 KB: los largos se guardan en trozos. */
function cacheGuardar_(clave, texto, seg) {
  const c = CacheService.getScriptCache(), T = 40000, n = Math.max(1, Math.ceil(texto.length / T)), o = {};
  for (let i = 0; i < n; i++) o[clave + ':' + i] = texto.slice(i * T, (i + 1) * T);
  o[clave] = String(n);
  c.putAll(o, seg);
}
function cacheLeer_(clave) {
  const c = CacheService.getScriptCache(), n = Number(c.get(clave) || 0);
  if (!n) return null;
  const ks = []; for (let i = 0; i < n; i++) ks.push(clave + ':' + i);
  const v = c.getAll(ks); let t = '';
  for (const k of ks) { if (v[k] == null) return null; t += v[k]; }
  return t;
}
function cacheBorrar_(clave) { CacheService.getScriptCache().remove(clave); }

/* ================= Registro horario ================= */

/** Devuelve la pestaña del registro horario; si no existe, la crea con sus columnas. */
function hojaRegistro_() {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(HOJA.reg);
  if (!h) {
    h = ss.insertSheet(HOJA.reg);
    h.getRange(1, 1, 1, 12).setValues([['ID', 'Año', 'ID empleado', 'Empleado', 'Fecha', 'Entrada', 'Salida', 'Horas', 'Origen', 'Nota', 'Registrado el', 'Modificado el']]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}
/** Horas entre dos horas "HH:mm" (si la salida es menor, se entiende que es al día siguiente). */
function horas_(entrada, salida) { let m = minutos_(entrada, salida); if (m < 0) m += 1440; return Math.round(m / 60 * 100) / 100; }
function leerRegistro_(desde) {
  hojaRegistro_();
  return tabla_(HOJA.reg).filas.map(f => ({ id: String(f[0]), emp: f[2], fecha: fecha_(f[4]), entrada: hora_(f[5]), salida: hora_(f[6]),
    horas: f[7] === '' ? null : Number(f[7]), origen: ORIGEN_REG_I[f[8]] || 'movil', nota: f[9] || '' }))
    .filter(r => r.fecha >= desde);
}

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

/** Deja programadas las tareas automáticas: revisión diaria de solicitudes y comprobación de actualizaciones cada hora. */
function instalarRecordatorio() {
  ScriptApp.getProjectTriggers().filter(t => ['recordatorioDiario', 'comprobarActualizacion'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('recordatorioDiario').timeBased().everyDays(1).atHour(CFG.HORA_RECORDATORIO).inTimezone('Europe/Madrid').create();
  ScriptApp.newTrigger('comprobarActualizacion').timeBased().everyHours(1).create();
  const msg = 'Tareas activadas: cada día hacia las ' + CFG.HORA_RECORDATORIO + ':00 se revisan las solicitudes pendientes, y cada hora se comprueba si hay una versión nueva del código.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

/* ================= Actualización automática desde GitHub ================= */

function apiScript_(metodo, ruta, cuerpo) {
  const r = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + ruta, {
    method: metodo, contentType: 'application/json', muteHttpExceptions: true,
    // se usa el proyecto de Firebase (que es tuyo) para la cuota de la API de Apps Script
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'x-goog-user-project': FIREBASE_PROYECTO },
    payload: cuerpo ? JSON.stringify(cuerpo) : undefined
  });
  const c = r.getResponseCode(), t = r.getContentText();
  if (c >= 300) {
    let m = t; try { m = JSON.parse(t).error.message; } catch (e) {}
    let pista = '';
    if (/has not been used in project|is disabled|SERVICE_DISABLED/i.test(m)) pista = 'Hay que activar la "Apps Script API" en el proyecto de Firebase: console.cloud.google.com/apis/library/script.googleapis.com?project=' + FIREBASE_PROYECTO + ' — ';
    else if (/User has not enabled the Apps Script API|usersettings/i.test(m)) pista = 'Activa la "API de Google Apps Script" en script.google.com/home/usersettings — ';
    throw new Error(pista + 'Error de Google: ' + m);
  }
  return t ? JSON.parse(t) : {};
}
function descargar_(archivo) {
  const r = UrlFetchApp.fetch(REPO_RAW + encodeURIComponent(archivo) + '?v=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('No se ha podido descargar ' + archivo + ' de GitHub (' + r.getResponseCode() + ').');
  return r.getContentText();
}

/** Cada hora: si en GitHub hay otra versión, la instala sin hacer nada a mano. */
function comprobarActualizacion() {
  try { avisarNominasNuevas_(); } catch (e) { console.warn('Nóminas: ' + e.message); }
  let codigo;
  try { codigo = descargar_('Codigo.gs'); } catch (e) { console.warn(e.message); return; }
  const m = codigo.match(/const VERSION_SERVIDOR = '([^']+)'/);
  if (!m || m[1] === VERSION_SERVIDOR) return;
  try {
    const r = aplicarActualizacion_(codigo);
    enviarAhora_('admin', 'App actualizada', 'La hoja tiene el código nuevo (versión ' + m[1] + ').' + (r.permisosNuevos ? ' Hay permisos nuevos: abre Apps Script y ejecuta "autorizar".' : ''));
  } catch (e) {
    console.warn('Actualización fallida: ' + e.message);
    const pr = PropertiesService.getScriptProperties();
    if (pr.getProperty('AVISO_ACTUALIZACION') !== m[1]) {   // avisa una sola vez por versión
      pr.setProperty('AVISO_ACTUALIZACION', m[1]);
      enviarAhora_('admin', 'No se ha podido actualizar la hoja', e.message);
    }
  }
}

/** También desde el menú App RRHH › Actualizar ahora. */
function actualizarAhora() {
  const ui = SpreadsheetApp.getUi();
  try {
    const codigo = descargar_('Codigo.gs'), m = codigo.match(/const VERSION_SERVIDOR = '([^']+)'/);
    if (m && m[1] === VERSION_SERVIDOR) { ui.alert('Ya tienes la última versión (' + VERSION_SERVIDOR + ').'); return; }
    const r = aplicarActualizacion_(codigo);
    ui.alert('Actualizado a la versión ' + (m ? m[1] : '?') + '.' + (r.permisosNuevos ? '\n\nHay permisos nuevos: en Apps Script elige "autorizar" y pulsa Ejecutar.' : ''));
  } catch (e) { ui.alert('No se ha podido actualizar: ' + e.message); }
}

function aplicarActualizacion_(codigo) {
  const manifiesto = descargar_('appsscript.json');
  JSON.parse(manifiesto);   // comprueba que está bien escrito
  const actual = apiScript_('get', '/content');
  const archivos = actual.files || [];
  const js = archivos.find(f => f.type === 'SERVER_JS'), man = archivos.find(f => f.type === 'JSON');
  if (!js || !man) throw new Error('No encuentro los archivos del proyecto.');
  const scopes = t => JSON.stringify((JSON.parse(t).oauthScopes || []).slice().sort());
  const permisosNuevos = scopes(man.source) !== scopes(manifiesto);
  js.source = codigo; man.source = manifiesto;
  apiScript_('put', '/content', { files: archivos });
  // con permisos nuevos no se publica todavía (la app dejaría de funcionar hasta autorizar): lo hace "autorizar"
  if (permisosNuevos) { PropertiesService.getScriptProperties().setProperty('PUBLICAR_PENDIENTE', '1'); return { permisosNuevos, version: null, implementaciones: 0 }; }
  return Object.assign({ permisosNuevos }, publicarVersion_());
}
function publicarVersion_() {
  const v = apiScript_('post', '/versions', { description: 'Actualización automática ' + hoy_() });
  const dep = (apiScript_('get', '/deployments').deployments || []).filter(d =>
    d.deploymentConfig && d.deploymentConfig.versionNumber && (d.entryPoints || []).some(e => e.entryPointType === 'WEB_APP'));
  dep.forEach(d => apiScript_('put', '/deployments/' + d.deploymentId, { deploymentConfig: {
    scriptId: ScriptApp.getScriptId(), versionNumber: v.versionNumber, manifestFileName: 'appsscript',
    description: 'Versión ' + v.versionNumber + ' (automática)' } }));
  return { version: v.versionNumber, implementaciones: dep.length };
}

/** Ejecútala una vez desde el editor (Ejecutar) para dar todos los permisos, también el de enviar avisos. */
function autorizar() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  UrlFetchApp.fetch('https://www.google.com');
  const pr = PropertiesService.getScriptProperties();
  try {
    const n = indiceNominas_(true).length;
    pr.setProperty('NOMINAS_LISTO', '1');
    avisarNominasNuevas_();   // deja apuntadas las que ya hay, para avisar solo de las nuevas
    console.log('Nóminas: ' + n + ' PDF encontrados en la carpeta.');
  } catch (e) { console.warn('No puedo leer la carpeta de nóminas: ' + e.message); }
  // si una actualización trajo permisos nuevos, ahora que están dados se publica
  if (pr.getProperty('PUBLICAR_PENDIENTE')) {
    try { publicarVersion_(); pr.deleteProperty('PUBLICAR_PENDIENTE'); console.log('Versión nueva publicada.'); } catch (e) { console.warn('No se ha podido publicar: ' + e.message); }
  }
  console.log('Permisos correctos.');
}

/* ================= Menú en la hoja ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('App RRHH')
    .addItem('Cambiar PIN del administrador', 'cambiarPinAdmin')
    .addItem('Activar tareas automáticas (recordatorio y actualizaciones)', 'instalarRecordatorio')
    .addItem('Actualizar ahora desde GitHub', 'actualizarAhora')
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
