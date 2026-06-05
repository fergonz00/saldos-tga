/**
 * saldos-tga · Apps Script de la planilla ESPEJO
 * Spreadsheet: https://docs.google.com/spreadsheets/d/19hKf6VaOsjGlk9s-biZtql5AW0oIV8oBlfqYGEfMH8I/edit
 *
 * Hojas:
 *   - "compras"  → espejo (IMPORTRANGE) del origen 1s7QlK99... Detalle de unidades
 *                  y filas especiales con los totales ya calculados.
 *   - "INFORME"  → espejo (IMPORTRANGE) de 1qIP4OT5... Saldos bancos / VW Credit.
 *
 * Publicación: Deploy → Web app · Execute as: Me · Who: Anyone (con token)
 */

const TOKEN = 'tga-saldos-K9Mx2P7vQ';
const SHEET_COMPRAS = 'compras';
const SHEET_INFORME = 'INFORME';

// Filas especiales (modelo en col I) que NO son unidades sino totales calculados.
const FILA_FLOOR_PLAN   = 'linea de credito floor plan';
const FILA_DEUDA_FP     = 'deuda floor plan';
const FILA_DISP_FP      = 'disponible floor plan';
const FILA_DISP_VW      = 'total disp real para pagar';

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (String(params.token || '').trim() !== TOKEN) return jsonResponse({ error: 'forbidden' });

  const tipo = String(params.tipo || 'all').toLowerCase();
  try {
    if (tipo === 'compras') return jsonResponse(getCompras());
    if (tipo === 'informe') return jsonResponse(getInforme());
    return jsonResponse({ compras: getCompras(), informe: getInforme() });
  } catch (err) {
    return jsonResponse({ error: String(err && err.message || err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =======================================================================
// COMPRAS — unidades + totales (todo viene de la hoja "compras")
// =======================================================================
function getCompras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_COMPRAS);
  if (!sh) throw new Error('No existe la hoja "' + SHEET_COMPRAS + '"');

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 18); // hasta R
  if (lastRow < 1) {
    return { unidades: [], floorPlan: 0, deudaFloorPlan: 0, disponibleFloorPlan: 0, disponibleVW: 0 };
  }

  // getDisplayValues = strings tal cual se ven en la planilla (respeta el formato de mes/preventa).
  // getValues       = tipos crudos (numbers para importes).
  const range = sh.getRange(1, 1, lastRow, lastCol);
  const display = range.getDisplayValues();
  const raw     = range.getValues();

  // El detalle arranca después del comparativo histórico. Busco el header por col C = "preventa".
  let headerRow = -1;
  for (let i = 0; i < display.length; i++) {
    const c = String(display[i][2] || '').toLowerCase().trim();
    if (c === 'preventa') { headerRow = i; break; }
  }
  if (headerRow < 0) throw new Error('No encontré el header del detalle (col C = "preventa")');

  const unidades = [];
  let floorPlan = 0, deudaFloorPlan = 0, disponibleFloorPlan = 0, disponibleVW = 0;

  for (let i = headerRow + 1; i < display.length; i++) {
    const drow = display[i];
    const rrow = raw[i];
    const modelo = String(drow[8] || '').trim();              // col I
    const modeloKey = modelo.toLowerCase();

    // Filas especiales de totales (no son unidades)
    if (modeloKey === FILA_FLOOR_PLAN)   { floorPlan           = toNumber(rrow[10]); continue; }
    if (modeloKey === FILA_DEUDA_FP)     { deudaFloorPlan      = toNumber(rrow[10]); continue; }
    if (modeloKey === FILA_DISP_FP)      { disponibleFloorPlan = toNumber(rrow[10]); continue; }
    if (modeloKey === FILA_DISP_VW)      { disponibleVW        = toNumber(rrow[10]); continue; }

    // Saltar filas claramente vacías
    const preventa = String(drow[2] || '').trim();
    const serie    = String(drow[7] || '').trim();
    if (!preventa && !serie && !modelo) continue;

    const importeK  = toNumber(rrow[10]);              // K — vacía cuando se paga
    const totalR    = toNumber(rrow[17]);              // R — total factura (siempre)
    const fechaPago = String(drow[14] || '').trim();   // O
    const impaga    = importeK > 0 && !fechaPago;

    unidades.push({
      idx:           toNumber(rrow[0]),                // A
      mes:           String(drow[1] || '').trim(),     // B "noviembre-25"
      preventa:      preventa,                         // C "8654/3"
      fechaVta:      String(drow[3] || '').trim(),     // D
      clienteCancelo:String(drow[4] || '').trim(),     // E
      mesPat:        String(drow[5] || '').trim(),     // F
      localidad:     String(drow[6] || '').trim(),     // G
      serie:         serie,                            // H
      modelo:        modelo,                           // I
      fc:            String(drow[9] || '').trim(),     // J
      importe:       importeK,                         // K
      color:         String(drow[11] || '').trim(),    // L
      fechaFc:       String(drow[12] || '').trim(),    // M
      vence:         String(drow[13] || '').trim(),    // N
      fechaPago:     fechaPago,                        // O
      fechaCertif:   String(drow[16] || '').trim(),    // Q
      total:         totalR,                           // R
      impaga:        impaga,
    });
  }

  return {
    unidades:            unidades,
    floorPlan:           floorPlan,
    deudaFloorPlan:      deudaFloorPlan,
    disponibleFloorPlan: disponibleFloorPlan,
    disponibleVW:        disponibleVW,
    updatedAt:           new Date().toISOString(),
  };
}

// =======================================================================
// INFORME — Bancos / disponibilidades / Personal
// =======================================================================
// Estructura real de la hoja INFORME (verificada con la data cruda):
//   - Col B (index 1): concepto (texto)
//   - Col C (index 2): valor primario (Libre Disp en DISP FUTURAS, monto en DISPONIBILIDADES)
//   - Col D (index 3): valor secundario (Técnico en DISP FUTURAS, fecha de venc.
//            en DISPONIBILIDADES, monto en SALDOS DISP VW / DEUDAS EN FABRICA)
//   - Col E (index 4): Personal — texto (NOVEDADES DEL PERSONAL / Ausentes /
//            Tarde / nombres)
//   - Col F (index 5): Personal — fechas/horas asociadas
//
// Cortes:
//   - Saldos: dejar de leer cuando el concepto contiene "rendimiento" (Rendimiento diario Fondos)
//   - Personal: cortar al detectar "proveedor" o "factura"
//
// Cada fila devuelve { tipo, concepto, d, dStr, e, eStr, esFechaE } para que el
// frontend decida qué mostrar.
function getInforme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_INFORME);
  if (!sh) throw new Error('No existe la hoja "' + SHEET_INFORME + '"');

  const lastRow = Math.max(sh.getLastRow(), 80);
  const lastCol = Math.max(sh.getLastColumn(), 8);
  if (lastRow < 3) {
    return { rows: [], disponibleVW: 0, personal: [], updatedAt: new Date().toISOString() };
  }

  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const raw     = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // ===== SALDOS (col C / D / E, fila 3 en adelante) =====
  const rows = [];
  let disponibleVW = 0;
  let lastWasSep = true;
  const RE_FECHA = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/;

  // i=1 → fila 2 (donde empieza "DISPONIBILIDADES $" como header)
  for (let i = 1; i < display.length; i++) {
    const concepto = String(display[i][1] || '').trim();   // B
    const dStr     = String(display[i][2] || '').trim();   // C
    const eStr     = String(display[i][3] || '').trim();   // D
    const dNum     = toNumber(raw[i][2]);
    const eNum     = toNumber(raw[i][3]);

    // Stop a "rendimiento" si aparece en cualquier columna leída
    const blob = (concepto + ' ' + dStr + ' ' + eStr).toLowerCase();
    if (blob.indexOf('rendimiento') >= 0) break;

    // Capturar disponible VW
    if (!disponibleVW && /total\s+disp\b/i.test(concepto)) disponibleVW = dNum || eNum;
    if (!disponibleVW && /saldos\s+disponibles\s+vw/i.test(concepto)) disponibleVW = dNum || eNum;

    // Fila totalmente vacía → separador
    if (!concepto && !dStr && !eStr) {
      if (!lastWasSep) { rows.push({ tipo: 'sep' }); lastWasSep = true; }
      continue;
    }

    // Fila con concepto vacío pero con valor:
    //   En la planilla, los TOTALES de DISP FUTURAS y similares ponen el
    //   valor "Técnico" en una fila aparte. Lo mergeo al row anterior si era
    //   un total al que le falta valor E.
    if (!concepto && (dStr || eStr)) {
      if (rows.length) {
        const prev = rows[rows.length - 1];
        if (prev.tipo === 'total' && prev.e === 0 && eNum) {
          prev.e    = eNum;
          prev.eStr = eStr;
        }
      }
      continue;
    }
    lastWasSep = false;

    // Clasificar: whitelist de headers de sección (solo los nombres conocidos)
    let tipo = 'row';
    if (_esHeaderSeccion(concepto))                  tipo = 'header';
    else if (/^(total|subtotal)\b/i.test(concepto))  tipo = 'total';

    rows.push({
      tipo:     tipo,
      concepto: concepto,
      d:        dNum,
      dStr:     dStr,
      e:        eNum,
      eStr:     eStr,
      esFechaE: RE_FECHA.test(eStr),
    });
  }
  while (rows.length && rows[rows.length - 1].tipo === 'sep') rows.pop();

  // ===== FONDOS — Rendimiento diario =====
  //   Col B (1): fecha del análisis (ej "11-may")
  //   Col C (2): total acumulado en el fondo
  //   Col D (3): ganancia del día ($)
  //   Col E (4): % diario
  //   Col F (5): % anual
  //   Col G (6): % mensual
  // Arranca al detectar "Rendimiento diario Fondos" en cualquier columna.
  // Termina al detectar "acumulado" (sección "Rendimiento acumulado" más abajo).
  const fondos = [];
  let inFondos = false;
  for (let i = 0; i < display.length; i++) {
    const blob = display[i].join(' ').toLowerCase();
    if (!inFondos) {
      if (blob.indexOf('rendimiento diario') >= 0) inFondos = true;
      continue;
    }
    if (blob.indexOf('acumulado') >= 0) break;

    const fecha       = String(display[i][1] || '').trim();
    const totalStr    = String(display[i][2] || '').trim();
    const gananciaStr = String(display[i][3] || '').trim();
    const diarioStr   = String(display[i][4] || '').trim();
    const anualStr    = String(display[i][5] || '').trim();
    const mensualStr  = String(display[i][6] || '').trim();
    const totalNum    = toNumber(raw[i][2]);
    const gananciaNum = toNumber(raw[i][3]);

    // Filas claramente vacías
    if (!fecha && !totalStr && !gananciaStr && !diarioStr) continue;

    fondos.push({
      fecha:       fecha,
      total:       totalNum,
      totalStr:    totalStr,
      ganancia:    gananciaNum,
      gananciaStr: gananciaStr,
      diario:      diarioStr,
      anual:       anualStr,
      mensual:     mensualStr,
    });
  }

  // ===== PERSONAL (col E texto + col F fecha/hora, fila 3 en adelante) =====
  const personal = [];
  for (let i = 2; i < display.length; i++) {
    const txt = String(display[i][4] || '').trim();   // E
    const dat = String(display[i][5] || '').trim();   // F
    if (!txt) {
      if (personal.length && personal[personal.length - 1].tipo !== 'sep') {
        personal.push({ tipo: 'sep' });
      }
      continue;
    }
    // Cortar al entrar a PROVEEDORES o FACTURAS
    if (/proveedor/i.test(txt) || /^factura/i.test(txt)) break;
    personal.push(_clasificarFilaPersonal(txt, dat));
  }
  while (personal.length && personal[personal.length - 1].tipo === 'sep') personal.pop();

  return {
    rows:         rows,
    disponibleVW: disponibleVW,
    fondos:       fondos,
    personal:     personal,
    updatedAt:    new Date().toISOString(),
  };
}

// Headers de sección en Bancos. Sólo los nombres conocidos califican —
// así "AFIP - SALDO IVA" no se confunde con header sólo por estar en mayúsculas.
function _esHeaderSeccion(concepto) {
  const u = concepto.trim().toUpperCase();
  return /^DISPONIBILIDADES/.test(u) ||
         /^DISP\s+FUTURAS/.test(u) ||
         /^DETALLE\s+DE\s+CHEQDIF/.test(u) ||
         /^SALDOS\s+DISPONIBLES\s+VW/.test(u) ||
         /^DEUDAS\s+EN\s+FABRICA/.test(u);
}

function _clasificarFilaPersonal(texto, fecha) {
  if (/novedades\s+del\s+personal/i.test(texto)) {
    return { tipo: 'header', texto: texto };
  }
  // Subheader corto: Ausentes / Tarde / Reincorp(orados) — sin números, sin paréntesis ni dos puntos
  const corto = texto.length <= 30;
  const sinDigitos = !/\d/.test(texto);
  const sinDelimiter = !/[:\(\)\-,]/.test(texto);
  if (corto && sinDigitos && sinDelimiter) {
    return { tipo: 'subhead', texto: texto };
  }
  return { tipo: 'item', texto: texto, fecha: fecha };
}

// =======================================================================
// HELPERS
// =======================================================================
function toNumber(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;
  if (v instanceof Date) return 0;   // fechas NO son números
  const s = String(v).replace(/[^\d.,-]/g, '');
  if (!s) return 0;
  const hasComma = s.indexOf(',') > -1;
  const hasDot   = s.indexOf('.') > -1;
  let n;
  if (hasComma && hasDot) {
    n = parseFloat(s.replace(/\./g, '').replace(',', '.'));
  } else if (hasComma) {
    n = parseFloat(s.replace(',', '.'));
  } else {
    n = parseFloat(s);
  }
  return isNaN(n) ? 0 : n;
}

// =======================================================================
// NOVEDADES · snapshot diario del personal general → Supabase (acumulación)
// =======================================================================
// Guarda las novedades del personal (Ausentes / Tarde) de la hoja INFORME en
// la tabla saldos_novedades de Supabase, para poder sacar reportes históricos.
// Idempotente: dedup_key evita duplicar en corridas repetidas.
//
// SETUP (una sola vez): correr instalarTriggerNovedades() desde el editor.
//   Asegurate que el proyecto tenga timezone Argentina (Configuración del proyecto).
const SUPABASE_URL_NOV  = 'https://wjfglsafgaltusmbnccl.supabase.co';
const SUPABASE_ANON_NOV = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndqZmdsc2FmZ2FsdHVzbWJuY2NsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM0MzM2OTksImV4cCI6MjA4OTAwOTY5OX0.OOwgyKDNQsbBaGDaL0OhJfc8eOsCClvvAPW0VFBKrOA';

function snapshotNovedades() {
  const inf = getInforme();
  const personal = (inf && inf.personal) || [];
  let seccion = '';
  const eventos = [];
  personal.forEach(function (r) {
    if (r.tipo === 'header')  { seccion = ''; return; }
    if (r.tipo === 'subhead') { seccion = String(r.texto || '').trim(); return; }
    if (r.tipo !== 'item') return;
    const texto = String(r.texto || '').trim();
    if (!texto) return;
    const tipo = _tipoNovedad(seccion);
    const fechaISO = _fechaNovedadISO(r.fecha);
    eventos.push({
      fecha: fechaISO, area: 'general', tipo: tipo, persona: texto,
      detalle: String(r.fecha || '').trim() || null, origen: 'sheet',
      dedup_key: [fechaISO, 'general', tipo, texto.toLowerCase()].join('|'),
    });
  });
  if (!eventos.length) return 0;
  const res = UrlFetchApp.fetch(SUPABASE_URL_NOV + '/rest/v1/saldos_novedades?on_conflict=dedup_key', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'apikey': SUPABASE_ANON_NOV,
      'Authorization': 'Bearer ' + SUPABASE_ANON_NOV,
      'Prefer': 'resolution=ignore-duplicates,return=minimal',
    },
    payload: JSON.stringify(eventos),
    muteHttpExceptions: true,
  });
  Logger.log('snapshotNovedades: ' + eventos.length + ' eventos · HTTP ' + res.getResponseCode());
  return eventos.length;
}

function _tipoNovedad(seccion) {
  const s = String(seccion || '').toLowerCase();
  if (s.indexOf('tarde') > -1) return 'llegada_tarde';
  if (/ausent|falt/.test(s)) return 'falta';
  return 'otro';
}

function _fechaNovedadISO(f) {
  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!f) return hoy;
  const s = String(f).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (m) {
    var d = ('0' + m[1]).slice(-2), mo = ('0' + m[2]).slice(-2);
    var y = m[3] ? (m[3].length === 2 ? '20' + m[3] : m[3]) : String(new Date().getFullYear());
    return y + '-' + mo + '-' + d;
  }
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) return s;
  return hoy;
}

// Correr UNA vez para programar el chequeo diario ~16:00 (hora del proyecto).
function instalarTriggerNovedades() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'snapshotNovedades') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('snapshotNovedades').timeBased().everyDays(1).atHour(16).create();
  Logger.log('Trigger diario instalado: snapshotNovedades ~16hs.');
}
