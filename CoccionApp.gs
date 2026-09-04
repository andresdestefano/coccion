/**
 * CoccionApp.gs — Backend de la App de Cocción Hammurabier
 * Versión B3 · 2026-09-04 — B3: api `recetaIA` (Gemini) para completar la receta desde texto libre. B2: columnas en formato texto (@) para que Sheets no convierta pasos/fechas en Date; lectura normaliza Date→string.
 *
 * Hoja: "Cocciones Hammurabier" con pestañas Lotes y Registros.
 * API (GET o POST): ?api=<funcion>&datos=<json urlencoded>
 *   lotes         → lista de lotes (abiertos primero, luego últimos 30 cerrados)
 *   loteCrear     {receta:{...}}                     → crea lote, devuelve id
 *   loteGet       {id}                               → lote + registros
 *   loteEditar    {id, receta:{...}}                 → reemplaza receta (solo ABIERTO)
 *   loteEstado    {id, estado:'ABIERTO'|'CERRADO'}   → cambia estado
 *   pasoRegistrar {id, paso, hecho, ts, operador, valores:{...}, nota, uid}
 *                                                    → registra evento; uid dedupe (idempotente)
 *   pasoLote      {id, eventos:[{...}]}              → varios eventos en una llamada (cola offline)
 *   setup                                            → crea hoja/pestañas si faltan (solo dueño)
 *   ping                                             → {ok:true, version}
 *   recetaIA      {texto}                            → {ok, receta:{...}} vía Gemini (GEMINI_API_KEY en Script Properties)
 *
 * Toda escritura bajo LockService. Registros es append-only: el estado actual de un paso
 * es su ÚLTIMO evento. Deshacer = evento con hecho=0.
 */

var SHEET_ID = ''; // vacío = script vinculado a la hoja (SpreadsheetApp.getActive)
var VERSION = 'B3';
var LOTES_COLS = ['id','creado','estado','nombre','estilo','tipo','fecha','operador','receta','inicio','fin'];
var REG_COLS = ['uid','id_lote','paso','ts','hecho','operador','valores','nota','recibido'];

function doGet(e)  { return api_(e); }
function doPost(e) { return api_(e); }

function api_(e) {
  var out;
  try {
    var p = (e && e.parameter) || {};
    var api = p.api || '';
    var raw = p.datos || (e && e.postData && e.postData.contents) || '{}';
    var d = {};
    try { d = JSON.parse(raw); } catch (err) { d = {}; }
    if (!api && d.api) { api = d.api; d = d.datos || {}; }
    switch (api) {
      case 'ping':          out = { ok: true, version: VERSION }; break;
      case 'lotes':         out = lotes_(); break;
      case 'loteCrear':     out = loteCrear_(d); break;
      case 'loteGet':       out = loteGet_(d); break;
      case 'loteEditar':    out = loteEditar_(d); break;
      case 'loteEstado':    out = loteEstado_(d); break;
      case 'pasoRegistrar': out = pasoLote_({ id: d.id, eventos: [d] }); break;
      case 'pasoLote':      out = pasoLote_(d); break;
      case 'setup':         out = setup_(); break;
      case 'recetaIA':      out = recetaIA_(d); break;
      default:              out = { ok: false, error: 'api no reconocida: ' + api };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- acceso a hoja ----------
function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}
function hoja_(nombre, cols) {
  var ss = ss_();
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.setFrozenRows(1);
  }
  textoCols_(sh, cols.length);
  return sh;
}
function setup_() {
  hoja_('Lotes', LOTES_COLS);
  hoja_('Registros', REG_COLS);
  return { ok: true, version: VERSION };
}
function filas_(sh, cols) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var res = [];
  for (var i = 0; i < vals.length; i++) {
    var o = { _row: i + 2 };
    for (var j = 0; j < cols.length; j++) o[cols[j]] = vals[i][j];
    res.push(o);
  }
  return res;
}
function textoCols_(sh, n) {
  // Formato texto en todas las columnas: evita que Sheets convierta '2.6.17' o '2026-09-03' en fechas.
  var key = 'txt_' + sh.getSheetId() + '_' + n;
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(key)) return;
  sh.getRange(1, 1, sh.getMaxRows(), n).setNumberFormat('@');
  props.setProperty(key, '1');
}
function escribir_(sh, row, col, filas) {
  var rg = sh.getRange(row, col, filas.length, filas[0].length);
  rg.setNumberFormat('@');
  rg.setValues(filas);
}
function str_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : v.toISOString();
  return String(v);
}
function json_(s, def) { try { return s ? JSON.parse(s) : def; } catch (e) { return def; } }
function now_() { return new Date().toISOString(); }
function nuevoId_() {
  var d = new Date();
  var y = d.getFullYear(), m = ('0' + (d.getMonth() + 1)).slice(-2), dd = ('0' + d.getDate()).slice(-2);
  return 'L' + y + m + dd + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
}
function conLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- lotes ----------
function loteFila_(o) {
  return {
    id: str_(o.id), creado: str_(o.creado), estado: str_(o.estado), nombre: str_(o.nombre),
    estilo: str_(o.estilo), tipo: str_(o.tipo), fecha: str_(o.fecha), operador: str_(o.operador),
    receta: json_(str_(o.receta), {}), inicio: str_(o.inicio), fin: str_(o.fin)
  };
}
function lotes_() {
  var sh = hoja_('Lotes', LOTES_COLS);
  var todos = filas_(sh, LOTES_COLS).map(loteFila_);
  var abiertos = todos.filter(function (l) { return l.estado === 'ABIERTO'; });
  var cerrados = todos.filter(function (l) { return l.estado !== 'ABIERTO'; });
  abiertos.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
  cerrados.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
  var lista = abiertos.concat(cerrados.slice(0, 30)).map(function (l) {
    return { id: l.id, estado: l.estado, nombre: l.nombre, estilo: l.estilo, tipo: l.tipo, fecha: l.fecha, operador: l.operador, inicio: l.inicio, fin: l.fin };
  });
  return { ok: true, lotes: lista };
}
function validarReceta_(r) {
  if (!r || typeof r !== 'object') throw new Error('receta faltante');
  var req = ['nombre', 'estilo', 'tipo', 'fecha', 'operador', 'vol_adh', 'agua_mash', 'agua_total', 't_mash', 't_mac', 't_hervor', 't_ferm'];
  var faltan = req.filter(function (k) { return r[k] === undefined || r[k] === null || String(r[k]).trim() === ''; });
  if (faltan.length) throw new Error('faltan campos de receta: ' + faltan.join(', '));
  if (r.tipo !== 'ale' && r.tipo !== 'lager') throw new Error('tipo debe ser ale o lager');
  return r;
}
function loteCrear_(d) {
  var r = validarReceta_(d.receta);
  return conLock_(function () {
    var sh = hoja_('Lotes', LOTES_COLS);
    var id = d.id && /^L\d{8}-[A-Z0-9]{4}$/.test(d.id) ? d.id : nuevoId_();
    var existentes = filas_(sh, LOTES_COLS);
    for (var i = 0; i < existentes.length; i++) if (str_(existentes[i].id) === id) return { ok: true, id: id, dup: true };
    var fila = [id, now_(), 'ABIERTO', str_(r.nombre), str_(r.estilo), str_(r.tipo), str_(r.fecha), str_(r.operador), JSON.stringify(r), '', ''];
    escribir_(sh, sh.getLastRow() + 1, 1, [fila]);
    return { ok: true, id: id };
  });
}
function buscarLote_(sh, id) {
  var todos = filas_(sh, LOTES_COLS);
  for (var i = 0; i < todos.length; i++) if (str_(todos[i].id) === str_(id)) return todos[i];
  return null;
}
function loteGet_(d) {
  if (!d.id) throw new Error('id faltante');
  var sh = hoja_('Lotes', LOTES_COLS);
  var o = buscarLote_(sh, d.id);
  if (!o) return { ok: false, error: 'lote no existe: ' + d.id };
  var regs = filas_(hoja_('Registros', REG_COLS), REG_COLS)
    .filter(function (r) { return str_(r.id_lote) === str_(d.id); })
    .map(function (r) {
      return { uid: str_(r.uid), paso: str_(r.paso), ts: str_(r.ts), hecho: Number(r.hecho) ? 1 : 0, operador: str_(r.operador), valores: json_(str_(r.valores), {}), nota: str_(r.nota) };
    });
  return { ok: true, lote: loteFila_(o), registros: regs };
}
function loteEditar_(d) {
  if (!d.id) throw new Error('id faltante');
  var r = validarReceta_(d.receta);
  return conLock_(function () {
    var sh = hoja_('Lotes', LOTES_COLS);
    var o = buscarLote_(sh, d.id);
    if (!o) return { ok: false, error: 'lote no existe' };
    if (str_(o.estado) !== 'ABIERTO') return { ok: false, error: 'lote cerrado: no se edita' };
    escribir_(sh, o._row, 4, [[str_(r.nombre), str_(r.estilo), str_(r.tipo), str_(r.fecha), str_(r.operador), JSON.stringify(r)]]);
    return { ok: true };
  });
}
function loteEstado_(d) {
  if (!d.id) throw new Error('id faltante');
  if (d.estado !== 'ABIERTO' && d.estado !== 'CERRADO') throw new Error('estado inválido');
  return conLock_(function () {
    var sh = hoja_('Lotes', LOTES_COLS);
    var o = buscarLote_(sh, d.id);
    if (!o) return { ok: false, error: 'lote no existe' };
    escribir_(sh, o._row, 3, [[d.estado]]);
    if (d.estado === 'CERRADO') escribir_(sh, o._row, 11, [[now_()]]);
    return { ok: true };
  });
}

// ---------- registros ----------
function pasoLote_(d) {
  if (!d.id) throw new Error('id faltante');
  var evs = d.eventos || [];
  if (!evs.length) return { ok: true, aplicados: 0, dups: 0 };
  return conLock_(function () {
    var shL = hoja_('Lotes', LOTES_COLS);
    var o = buscarLote_(shL, d.id);
    if (!o) return { ok: false, error: 'lote no existe: ' + d.id };
    if (str_(o.estado) !== 'ABIERTO') return { ok: false, error: 'lote cerrado: no acepta registros' };
    var shR = hoja_('Registros', REG_COLS);
    var uids = {};
    var last = shR.getLastRow();
    if (last >= 2) {
      var col = shR.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < col.length; i++) uids[str_(col[i][0])] = 1;
    }
    var filas = [], dups = 0, rech = [];
    var recibido = now_();
    for (var j = 0; j < evs.length; j++) {
      var ev = evs[j] || {};
      if (!ev.paso) { rech.push('sin paso'); continue; }
      var uid = str_(ev.uid) || (str_(d.id) + '|' + str_(ev.paso) + '|' + str_(ev.ts || recibido));
      if (uids[uid]) { dups++; continue; }
      uids[uid] = 1;
      filas.push([uid, str_(d.id), str_(ev.paso), str_(ev.ts || recibido), ev.hecho ? 1 : 0, str_(ev.operador), JSON.stringify(ev.valores || {}), str_(ev.nota), recibido]);
    }
    if (filas.length) escribir_(shR, shR.getLastRow() + 1, 1, filas);
    // marcar inicio del lote con el primer registro
    if (filas.length && !str_(o.inicio)) escribir_(shL, o._row, 10, [[recibido]]);
    return { ok: true, aplicados: filas.length, dups: dups, rechazados: rech };
  });
}

// ---------- receta por IA (Gemini) ----------
var RECETA_IA_CAMPOS = [
  ['estilo', 'string', 'nombre de la cerveza / estilo (ej. "Lalemana Helles Lager")'],
  ['tipo', 'string', '"ale" o "lager" según la levadura o el estilo'],
  ['vol_adh', 'number', 'volumen post-hervor / al fermentador en litros (NO el pre-hervor)'],
  ['og', 'number', 'gravedad original objetivo en °Plato'],
  ['vol_pre', 'number', 'volumen de recolección pre-hervor en litros'],
  ['grist', 'string', 'maltas con kg, separadas por "; ", y total entre paréntesis'],
  ['molienda', 'string', 'separación de rodillos / ajuste de molino si se indica'],
  ['agua_mash', 'number', 'agua de maceración en litros'],
  ['agua_total', 'number', 'agua tratada en HLT en litros = maceración + sparge + precalentamiento; NO incluir enjuague pre-cocina'],
  ['t_mash', 'number', 'temperatura de maceración en °C'],
  ['t_mac', 'number', 'tiempo de sacarificación/maceración en minutos (sin vorlauf)'],
  ['ph_mash', 'number', 'pH objetivo de maceración medido a 20 °C'],
  ['sales', 'string', 'sales al agua con cantidad, separadas por "; " (ej. "CaSO4·2H2O 40 g; CaCl2·2H2O 260 g")'],
  ['ascorbico', 'string', 'ácido ascórbico con cantidad (ej. "100 g")'],
  ['fosforico_mash', 'string', 'ácido fosfórico para la fracción de maceración, con cantidad en mL'],
  ['fosforico_sparge', 'string', 'ácido fosfórico para la fracción de sparge, con cantidad en mL'],
  ['t_hervor', 'number', 'duración del hervor en minutos'],
  ['adiciones', 'string', 'adiciones del hervor, una por línea, formato "faltan N min: producto — cantidad" (N = minutos que faltan para el fin del hervor; flameout = 0)'],
  ['t_ferm', 'number', 'temperatura de knock out / inicio de fermentación en °C'],
  ['t_clt', 'number', 'temperatura del CLT (agua de enfriamiento) en °C si se indica'],
  ['levadura', 'string', 'cepa y cantidad de levadura'],
  ['adiciones_fv', 'string', 'otras adiciones al fermentador (nutrientes, enzimas) con cantidad, separadas por "; "'],
  ['o2', 'string', '"Sí" o "No" si el documento dice si se oxigena; vacío si no lo dice'],
  ['fv', 'string', 'fermentador asignado (ej. "FV3") si se indica'],
  ['notas', 'string', 'datos operativos útiles sin campo propio: temperatura y pH de sparge, criterios de corte de lauterado, set del HLT, whirlpool, desviaciones registradas. Máximo 400 caracteres']
];
function recetaIA_(d) {
  var texto = str_(d && d.texto).trim();
  if (!texto) throw new Error('texto faltante');
  if (texto.length > 60000) texto = texto.slice(0, 60000);
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) return { ok: false, error: 'falta GEMINI_API_KEY en Propiedades del script' };
  var modelo = props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  var schemaProps = {}, desc = [];
  RECETA_IA_CAMPOS.forEach(function (c) { schemaProps[c[0]] = { type: c[1] === 'number' ? 'number' : 'string', nullable: true }; desc.push('- ' + c[0] + ' (' + c[1] + '): ' + c[2]); });
  var prompt = 'Sos el asistente técnico de una cervecería. Del siguiente documento de receta extraé los datos para la ficha de cocción. Reglas: usá exactamente las unidades pedidas (litros, °C, minutos, °Plato, pH a 20 °C); si un dato no está en el documento devolvé null, NUNCA lo inventes ni lo calcules salvo agua_total (suma de fracciones del HLT); no confundas volumen pre-hervor con post-hervor; respondé solo el JSON.\n\nCampos:\n' + desc.join('\n') + '\n\nDOCUMENTO:\n' + texto;
  var payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: { type: 'object', properties: schemaProps } }
  };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent';
  var resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(payload), muteHttpExceptions: true });
  var code = resp.getResponseCode(); var body = resp.getContentText();
  if (code !== 200) return { ok: false, error: 'Gemini HTTP ' + code + ': ' + body.slice(0, 300) };
  var j = json_(body, {});
  var txt = ''; try { txt = j.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join(''); } catch (e) { txt = ''; }
  var r = json_(txt.replace(/^```json\s*|```\s*$/g, ''), null);
  if (!r || typeof r !== 'object') return { ok: false, error: 'respuesta de Gemini no es JSON: ' + txt.slice(0, 200) };
  var out = {};
  RECETA_IA_CAMPOS.forEach(function (c) {
    var v = r[c[0]]; if (v === null || v === undefined || v === '') return;
    if (c[1] === 'number') { var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/(\d),(?=\d{3}\b)/g, '$1').replace(',', '.')); if (isNaN(n)) return; out[c[0]] = n; }
    else { out[c[0]] = String(v).trim(); }
  });
  if (out.tipo) out.tipo = /lager/i.test(out.tipo) ? 'lager' : 'ale';
  if (out.o2) out.o2 = /^s/i.test(out.o2) ? 'Sí' : 'No';
  return { ok: true, receta: out, modelo: modelo, usage: j.usageMetadata || null };
}
