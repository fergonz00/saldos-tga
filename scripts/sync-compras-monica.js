#!/usr/bin/env node
/**
 * sync-compras-monica.js
 *
 * Vuelca las facturas que Monica Marandola carga en su planilla "COMPRAS TGA"
 * (1s7QlK99..., pestana `compras`) a la tabla `compras_vw` de Supabase (wjfgl),
 * que es lo que muestra la solapa Compras VW de saldos.titogonzalez.online.
 *
 * Circuito (ver memoria project_saldos_compras_vw):
 *   1. La solapa Reparto crea sola la fila en compras_vw (updated_by='reparto')
 *      con serie/mes/modelo/color y TODO lo de factura en NULL.
 *   2. Monica anota la factura en su planilla.
 *   3. Este script copia importe/FC/vto/impuestos a compras_vw.
 *
 * Solo escribe filas que ya existen en compras_vw y que todavia NO tienen
 * factura cargada (importe_saldo NULL o 0). Nunca pisa una fila ya cargada:
 * si detecta diferencias contra la planilla las reporta y sigue.
 *
 * Uso:
 *   node sync-compras-monica.js            aplica los cambios
 *   node sync-compras-monica.js --dry      solo reporta, no escribe
 *   node sync-compras-monica.js --desde=2026-07-01   corta por fecha de FC
 */

const fs = require('fs');
const path = require('path');

const SSID = '1s7QlK99IiNYXZkNcpqEwMKHGtOqszw1vHKGmomzXVMg';
const TAB = 'compras';
const SB_REF = 'wjfglsafgaltusmbnccl';
const SECRETS = 'C:\\proyectos\\.secrets\\simulador-vwfs.env';
const ENV_CHAT = 'C:\\proyectos\\chat-tga\\.env.local';

const argv = process.argv.slice(2);
const DRY = argv.includes('--dry');
const DESDE = (argv.find(a => a.startsWith('--desde=')) || '--desde=2026-07-01').split('=')[1];
// Techo: las filas viejas de nov/dic-2025 de la planilla tienen el ano tipeado
// como 2026 (119-638xxx/119-640xxx, anteriores a las de enero-26). Sin este
// corte entran como si fueran futuras. 30 dias de margen alcanza de sobra.
const HASTA = new Date(Date.now() + 30 * 864e5).toISOString().slice(0, 10);

// Columnas de la pestana `compras` (0-indexed sobre el array de valores)
const C = {
  mes: 1, serie: 7, modelo: 8, fc: 9, importe: 10, color: 11,
  fechaFc: 12, vence: 13, fechaPago: 14, fechaCertif: 16,
  total: 17, iva: 18, neto: 19, percepIva: 20, iibbEr: 21, iibb: 22, iibbBsAs: 23, impInt: 24,
};

function readEnv(file) {
  const out = {};
  for (const line of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^([A-Z0-9_]+)=(.*)$/);
    if (m) out[m[1]] = m[2].trim();
  }
  return out;
}

const D = s => (typeof s === 'string' && /^\d{4}-\d\d-\d\dT/.test(s)) ? s.slice(0, 10) : null;
const N = x => (typeof x === 'number') ? x : null;

/**
 * Monica a veces tipea los importes con punto como separador de miles
 * ("151.341"), y Sheets lo guarda como decimal 151.341. Se detecta porque
 * el valor no es entero y es chico; el x1000 se valida despues con el cuadre.
 */
const fixMiles = v => (v != null && !Number.isInteger(v) && Math.abs(v) < 100000) ? Math.round(v * 1000) : v;

async function main() {
  const env = readEnv(SECRETS);
  const key = readEnv(ENV_CHAT).SUPABASE_SERVICE_ROLE_KEY;
  if (!env.PP_EXEC_URL || !env.PP_TOKEN) throw new Error('faltan PP_EXEC_URL/PP_TOKEN en ' + SECRETS);
  if (!key) throw new Error('falta SUPABASE_SERVICE_ROLE_KEY en ' + ENV_CHAT);

  // --- 1. planilla de Monica ---
  const url = `${env.PP_EXEC_URL}?panel=inspect&token=${env.PP_TOKEN}&id=${SSID}&tab=${TAB}`;
  const sheet = await (await fetch(url, { redirect: 'follow' })).json();
  if (!sheet.ok) throw new Error('no pude leer la planilla: ' + JSON.stringify(sheet).slice(0, 300));

  const filas = [];
  for (let i = 1; i < sheet.values.length; i++) {
    const r = sheet.values[i];
    const serie = String(r[C.serie] || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{8}$/.test(serie)) continue;
    const fechaFc = D(r[C.fechaFc]);
    if (!fechaFc || fechaFc < DESDE || fechaFc > HASTA) continue;
    filas.push({
      fila: i + 1, serie, modelo: r[C.modelo], fechaFc,
      fc: (r[C.fc] === '' || r[C.fc] == null) ? null : String(r[C.fc]).trim(),
      vence: D(r[C.vence]), fechaPago: D(r[C.fechaPago]), fechaCertif: D(r[C.fechaCertif]),
      total: N(r[C.total]), iva: N(r[C.iva]), neto: N(r[C.neto]),
      percepIva: N(r[C.percepIva]), iibbEr: N(r[C.iibbEr]),
      iibb: fixMiles(N(r[C.iibb])), iibbRaw: N(r[C.iibb]),
      iibbBsAs: N(r[C.iibbBsAs]), impInt: N(r[C.impInt]),
    });
  }

  // --- 2. portal ---
  const sb = async (p, init) => {
    const res = await fetch(`https://${SB_REF}.supabase.co/rest/v1/${p}`, {
      ...init,
      headers: { apikey: key, Authorization: `Bearer ${key}`, 'Content-Type': 'application/json', ...(init && init.headers) },
    });
    if (!res.ok) throw new Error(`supabase ${res.status}: ${await res.text()}`);
    const t = await res.text();
    return t ? JSON.parse(t) : null;
  };
  const db = await sb('compras_vw?select=*');
  const byS = Object.fromEntries(db.map(r => [r.serie, r]));

  // --- 3. clasificar ---
  const nuevas = [], yaCargadas = [], sinFila = [], sinFactura = [], alertas = [];

  for (const f of filas) {
    const p = byS[f.serie];
    if (!p) { sinFila.push(f); continue; }
    if (!f.total) { sinFactura.push(f); continue; }

    const suma = (f.neto || 0) + (f.iva || 0) + (f.percepIva || 0) + (f.iibbEr || 0) +
                 (f.iibb || 0) + (f.iibbBsAs || 0) + (f.impInt || 0);
    const dif = Math.round(f.total - suma);
    if (Math.abs(dif) > 2) alertas.push({ ...f, dif });

    const patch = {
      fc_numero: f.fc, fecha_fc: f.fechaFc, vence: f.vence,
      importe_saldo: f.total, neto: f.neto, iva_monto: f.iva,
      percep_iva: f.percepIva ?? 0, iibb_er: f.iibbEr, iibb_caba: f.iibb,
      iibb_bsas: f.iibbBsAs ?? 0, imp_internos: f.impInt ?? 0,
      updated_by: 'claude-sync-monica',
    };
    if (f.fechaPago) patch.fecha_pago_vw = f.fechaPago;
    if (f.fechaCertif) patch.fecha_certif = f.fechaCertif;

    if (p.importe_saldo == null || p.importe_saldo === 0) { nuevas.push({ f, patch }); continue; }

    // ya cargada: solo reportar diferencias, no pisar
    const eq = (a, b) => (a == null && b == null) || (a != null && b != null &&
      (typeof a === 'number' || typeof b === 'number' ? Math.abs(Number(a) - Number(b)) < 0.5 : String(a) === String(b)));
    const difs = Object.entries(patch)
      .filter(([k, v]) => k !== 'updated_by' && v != null && !eq(v, p[k]))
      .map(([k, v]) => `${k}: portal=${JSON.stringify(p[k])} planilla=${JSON.stringify(v)}`);
    if (difs.length) yaCargadas.push({ f, difs });
  }

  // --- 4. aplicar ---
  const log = [];
  const say = s => { console.log(s); log.push(s); };

  say(`[${new Date().toISOString()}] sync compras Monica -> saldos  (FC entre ${DESDE} y ${HASTA}${DRY ? ', DRY RUN' : ''})`);
  say(`planilla: ${filas.length} filas | portal: ${db.length} filas`);

  if (nuevas.length === 0) say('\nSin facturas nuevas para volcar.');
  else {
    say(`\n--- ${nuevas.length} factura(s) nueva(s) a cargar ---`);
    for (const { f, patch } of nuevas) {
      const linea = `${f.serie}  FC ${f.fc}  ${f.fechaFc}  vto ${f.vence}  $${(f.total || 0).toLocaleString('es-AR')}  ${f.modelo}`;
      if (DRY) { say('  [dry] ' + linea); continue; }
      await sb(`compras_vw?serie=eq.${encodeURIComponent(f.serie)}`, {
        method: 'PATCH', headers: { Prefer: 'return=minimal' }, body: JSON.stringify(patch),
      });
      say('  OK  ' + linea);
    }
    const tot = nuevas.reduce((a, x) => a + (x.f.total || 0), 0);
    say(`  TOTAL cargado: $${tot.toLocaleString('es-AR')}`);
  }

  if (sinFactura.length) {
    say(`\n--- ${sinFactura.length} en la planilla sin importe todavia (Monica no las facturo) ---`);
    sinFactura.forEach(f => say(`  ${f.serie}  fila ${f.fila}  ${f.modelo}`));
  }
  const pendientes = db.filter(r => (r.importe_saldo == null || r.importe_saldo === 0) &&
    r.updated_by === 'reparto' && !filas.some(f => f.serie === r.serie && f.total));
  if (pendientes.length) {
    say(`\n--- ${pendientes.length} en el portal esperando factura (no estan aun en la planilla) ---`);
    pendientes.forEach(r => say(`  ${r.serie}  ${r.mes}  ${r.modelo_valeria}`));
  }
  if (sinFila.length) {
    say(`\n--- ${sinFila.length} en la planilla que NO tienen fila en compras_vw ---`);
    sinFila.forEach(f => say(`  ${f.serie}  fila ${f.fila}  ${f.fechaFc}  ${f.modelo}`));
  }
  if (alertas.length) {
    say(`\n--- ${alertas.length} DESCUADRE(S) en la planilla (neto+IVA+percepciones != total) ---`);
    alertas.forEach(f => say(`  ${f.serie}  fila ${f.fila}  dif $${f.dif.toLocaleString('es-AR')}  (total ${f.total}, neto ${f.neto}, IVA ${f.iva}, ER ${f.iibbEr}, IIBB ${f.iibb}, BsAs ${f.iibbBsAs})`));
  }
  if (yaCargadas.length) {
    say(`\n--- ${yaCargadas.length} fila(s) ya cargada(s) con diferencias vs planilla (NO se pisaron) ---`);
    yaCargadas.forEach(({ f, difs }) => { say(`  ${f.serie} (fila ${f.fila})`); difs.forEach(d => say('     ' + d)); });
  }

  const dir = path.join(__dirname, 'logs');
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, `sync-${new Date().toISOString().slice(0, 10)}.log`), log.join('\n') + '\n');
}

main().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
