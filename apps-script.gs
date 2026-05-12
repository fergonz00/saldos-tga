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
// Estructura de la hoja INFORME (columnas C concepto / D valor):
//   - DISPONIBILIDADES $       (header)
//     Santander Cta Cte, Banco Galicia, Cheques, Efectivo, PFs, Fondos, USD, etc.
//     Subtotales: "Total Fondos de Inversion", "Total u$s"
//     Filas especiales: "Depósitos judiciales ...", "Stock de Planes ..."
//   - DISP FUTURAS              (header)
//     AFIP, IIBB CABA / Pcia / ER, Imp DB/CR, Cheques Diferidos, Total
//   - DETALLE DE CHEQDIF        (header)
//     Del 1 al 10, 11 al 20, 21 al 31, Meses siguiente
//   - SALDOS DISPONIBLES VW     (header con valor)  ← este es el disponible-VW
//     Saldo Disponible VW Credit x6
//     TOTAL DISP
//   - DEUDAS EN FABRICA         (header)
//     VW Arg y Financiera (unidades), TOTAL DEUDA VW Unidades, Total Deuda en VW - Disp
//   - (luego viene "Rendimiento diario Fondos" → SE CORTA acá)
//
// Para personal (col E):
//   - NOVEDADES DEL PERSONAL    (header), subhead Reincorp
//   - Ausentes (subheader) + items con fecha
//   - Tarde (subheader) + items con hora
//   - (luego viene "PROVEEDORES" → SE CORTA acá)
//
// Cada fila se devuelve con un `tipo` para que el frontend la pinte:
//   - 'header'   → encabezado de sección, todo en mayúsculas, sin número
//   - 'subhead'  → subencabezado (Ausentes / Tarde / Reincorp ...)
//   - 'total'    → fila que arranca con "Total" / "TOTAL" / "Subtotal"
//   - 'row'      → fila normal concepto + valor
//   - 'sep'      → fila en blanco (separador visual)
function getInforme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_INFORME);
  if (!sh) throw new Error('No existe la hoja "' + SHEET_INFORME + '"');

  const lastRow = Math.max(sh.getLastRow(), 80);
  const lastCol = Math.max(sh.getLastColumn(), 6);
  if (lastRow < 3) {
    return { rows: [], disponibleVW: 0, personal: [], updatedAt: new Date().toISOString() };
  }

  const display = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const raw     = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // ===== SALDOS (col C / D, fila 3 en adelante) =====
  const rows = [];
  let disponibleVW = 0;
  let lastWasSep = true;
  for (let i = 2; i < display.length; i++) {
    const concepto = String(display[i][2] || '').trim();
    const valorStr = String(display[i][3] || '').trim();
    const valorNum = toNumber(raw[i][3]);
    const lower    = concepto.toLowerCase();

    // Captura D47 = disponible VW (también puede venir del header SALDOS DISPONIBLES VW)
    if (i === 46) disponibleVW = valorNum;

    // Cortar al entrar en la sección "rendimiento diario de fondos"
    if (lower.indexOf('rendimiento') >= 0) break;

    // Fila vacía → separador
    if (!concepto && !valorStr) {
      if (!lastWasSep) { rows.push({ tipo: 'sep' }); lastWasSep = true; }
      continue;
    }
    lastWasSep = false;

    rows.push(_clasificarFilaSaldo(concepto, valorStr, valorNum));
  }
  while (rows.length && rows[rows.length - 1].tipo === 'sep') rows.pop();

  // Si el header "SALDOS DISPONIBLES VW" tiene su propio valor lo usamos
  // como fallback en caso de que D47 no haya capturado nada.
  if (!disponibleVW) {
    const found = rows.find(r => r.tipo === 'header' && /saldos\s+disponibles\s+vw/i.test(r.concepto || ''));
    if (found && found.valor) disponibleVW = found.valor;
  }

  // ===== PERSONAL (col E, fila 3 en adelante) =====
  const personal = [];
  for (let i = 2; i < display.length; i++) {
    const txt = String(display[i][4] || '').trim();
    if (!txt) {
      // Mantener un solo separador entre bloques
      if (personal.length && personal[personal.length - 1].tipo !== 'sep') {
        personal.push({ tipo: 'sep' });
      }
      continue;
    }
    // Sección "PROVEEDORES" o "FACTURAS" → corte
    if (/proveedor/i.test(txt) || /^factura/i.test(txt)) break;
    personal.push(_clasificarFilaPersonal(txt));
  }
  while (personal.length && personal[personal.length - 1].tipo === 'sep') personal.pop();

  return {
    rows:         rows,
    disponibleVW: disponibleVW,
    personal:     personal,
    updatedAt:    new Date().toISOString(),
  };
}

// Clasifica una fila de saldos en header/total/row.
function _clasificarFilaSaldo(concepto, valorStr, valorNum) {
  const letras = concepto.replace(/[^a-záéíóúñü]/gi, '');
  const esTodoMayus = letras.length >= 3 && letras === letras.toUpperCase();
  const empiezaTotal = /^(total|subtotal)\b/i.test(concepto);
  const tieneValor = valorNum !== 0 || (valorStr && valorStr !== '$ -' && valorStr !== '-');

  // Header de sección: texto todo en mayúsculas y SIN valor (DISPONIBILIDADES, DISP FUTURAS, etc.)
  // Excepción: "SALDOS DISPONIBLES VW" tiene valor pero queremos mostrarlo como header destacado.
  if (esTodoMayus && (!tieneValor || /saldos\s+disponibles\s+vw/i.test(concepto) || /deudas\s+en\s+fabrica/i.test(concepto))) {
    return { tipo: 'header', concepto, valor: valorNum, valorMostrar: valorStr };
  }
  if (empiezaTotal) {
    return { tipo: 'total', concepto, valor: valorNum, valorMostrar: valorStr };
  }
  return { tipo: 'row', concepto, valor: valorNum, valorMostrar: valorStr };
}

// Clasifica una fila de personal en header/subhead/item.
function _clasificarFilaPersonal(texto) {
  // Header grande (NOVEDADES DEL PERSONAL...)
  if (/novedades\s+del\s+personal/i.test(texto)) {
    return { tipo: 'header', texto };
  }
  // Subheader corto: Ausentes / Tarde / Reincorp(orados) — texto corto sin números/dos puntos
  const corto = texto.length <= 30;
  const sinDigitos = !/\d/.test(texto);
  const sinDelimiter = !/[:\(\)\-,]/.test(texto);
  if (corto && sinDigitos && sinDelimiter) {
    return { tipo: 'subhead', texto };
  }
  return { tipo: 'item', texto };
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
