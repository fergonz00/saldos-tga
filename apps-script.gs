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
// Lee la hoja "INFORME" y devuelve:
//   - saldos: [{ concepto, valor, valorMostrar }]   (col C/D desde fila 3
//             hasta detectar la sección "rendimiento")
//   - disponibleVW: número en D47 (disponible para aplicar en VW)
//   - personal: [{ texto }]                          (col E desde fila 3
//             hasta detectar la sección "factura")
function getInforme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_INFORME);
  if (!sh) throw new Error('No existe la hoja "' + SHEET_INFORME + '"');

  const lastRow = Math.max(sh.getLastRow(), 80);
  const lastCol = Math.max(sh.getLastColumn(), 6);
  if (lastRow < 3) {
    return { saldos: [], disponibleVW: 0, personal: [], updatedAt: new Date().toISOString() };
  }

  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const raw     = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // SALDOS / DISPONIBILIDADES (col C/D desde fila 3)
  const saldos = [];
  let disponibleVW = 0;
  let lastWasSep = true; // evitar separadores duplicados al inicio
  for (let i = 2; i < display.length; i++) {            // i=2 → fila 3
    const concepto = String(display[i][2] || '').trim(); // col C
    const valorStr = String(display[i][3] || '').trim(); // col D
    const valorNum = toNumber(raw[i][3]);
    const lower    = concepto.toLowerCase();

    // Captura especial: D47 = disponible VW (i=46 → fila 47)
    if (i === 46) disponibleVW = valorNum;

    // Cortar cuando entramos en "rendimiento diario de fondos"
    if (lower.indexOf('rendimiento') >= 0) break;

    // Fila completamente vacía → separador (para agrupar visualmente)
    if (!concepto && !valorStr) {
      if (!lastWasSep) { saldos.push({ separator: true }); lastWasSep = true; }
      continue;
    }
    lastWasSep = false;
    saldos.push({ concepto, valor: valorNum, valorMostrar: valorStr });
  }
  // Limpiar separador final
  while (saldos.length && saldos[saldos.length - 1].separator) saldos.pop();

  // PERSONAL (col E desde fila 3, hasta detectar "factura")
  const personal = [];
  for (let i = 2; i < display.length; i++) {
    const txt = String(display[i][4] || '').trim();    // col E
    if (!txt) continue;
    if (txt.toLowerCase().indexOf('factura') >= 0) break;
    personal.push({ texto: txt });
  }

  return {
    saldos:       saldos,
    disponibleVW: disponibleVW,
    personal:     personal,
    updatedAt:    new Date().toISOString(),
  };
}

// =======================================================================
// HELPERS
// =======================================================================
function toNumber(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;
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
