/**
 * Portal Logística · Indicadores de gestión
 * -----------------------------------------------------------------------
 * Función serverless de Vercel. Lee en SOLO LECTURA los Google Sheets de
 * las 4 apps y guarda lo propio del portal (usuarios, metas, cierres de
 * mes, registros excluidos y equivalencias) en un Sheet aparte.
 *
 * Variables de entorno (Vercel → Settings → Environment Variables):
 *   GOOGLE_SERVICE_ACCOUNT_EMAIL   correo de la cuenta de servicio
 *   GOOGLE_PRIVATE_KEY             llave privada (con los \n tal cual)
 *   SHEET_ID_PORTAL                Sheet "INDICADORES PORTAL" (lectura y escritura)
 *   SHEET_ID_CONDUCTORES           Sheet de la App Despachos (registro de conductores)
 *   SHEET_ID_ROLLOS                Sheet de la App Trazabilidad Rollos
 *   SHEET_ID_MONTACARGAS           Sheet del Checklist Montacargas
 *   SHEET_ID_IMPRESOS              Sheet de Control de Impresos, Insumos y Estibas
 *
 * Todo va en este único archivo (sin carpeta lib/) para evitar el error
 * "Cannot find module" que ya pasó en Despachos PT.
 */
const { google } = require('googleapis');

const ENV = process.env;
const IDS = {
  portal: ENV.SHEET_ID_PORTAL,
  cond: ENV.SHEET_ID_CONDUCTORES,
  rol: ENV.SHEET_ID_ROLLOS,
  mon: ENV.SHEET_ID_MONTACARGAS,
  imp: ENV.SHEET_ID_IMPRESOS
};

// ============================================================
// AUTENTICACIÓN (se reutiliza entre llamadas "tibias")
// ============================================================
let clientPromise = null;
function getSheets() {
  if (!clientPromise) {
    const key = (ENV.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
    const auth = new google.auth.JWT(ENV.GOOGLE_SERVICE_ACCOUNT_EMAIL, null, key, ['https://www.googleapis.com/auth/spreadsheets']);
    clientPromise = auth.authorize().then(() => google.sheets({ version: 'v4', auth })).catch((e) => { clientPromise = null; throw e; });
  }
  return clientPromise;
}

// ============================================================
// UTILIDADES DE LECTURA
// ============================================================
async function titles(sheets, id) {
  const r = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'sheets.properties.title' });
  return r.data.sheets.map((s) => s.properties.title);
}
async function readTabs(sheets, id, tabNames) {
  if (!tabNames.length) return {};
  const r = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: id, ranges: tabNames.map((t) => `'${t.replace(/'/g, "''")}'`),
    valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER'
  });
  const out = {};
  (r.data.valueRanges || []).forEach((vr, i) => { out[tabNames[i]] = vr.values || []; });
  return out;
}
const sinTilde = (s) => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
function findTab(list, ...names) {
  for (const n of names) { const f = list.find((t) => sinTilde(t) === sinTilde(n)); if (f) return f; }
  return null;
}
// Busca la fila de encabezados en las primeras 10 filas (Impresos la tiene en la fila 3)
function headerRow(rows, must) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const h = (rows[i] || []).map(sinTilde);
    if (must.every((m) => h.some((x) => x === sinTilde(m) || x.indexOf(sinTilde(m)) === 0))) return i;
  }
  return 0;
}
// Índice de la columna: primero nombre exacto, luego "empieza por"
function colIdx(header, ...names) {
  const h = header.map(sinTilde);
  for (const n of names) { const i = h.indexOf(sinTilde(n)); if (i >= 0) return i; }
  for (const n of names) { const i = h.findIndex((x) => x.indexOf(sinTilde(n)) === 0); if (i >= 0) return i; }
  return -1;
}
const cell = (r, i) => (i < 0 || !r ? '' : r[i] === undefined || r[i] === null ? '' : r[i]);

// ---- fechas, horas y números (pueden venir como número de Sheets o como TEXTO)
function serialToYMD(n) {
  const d = new Date(Math.round((Number(n) - 25569) * 864e5));
  return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
}
function fecha(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v > 20000 ? serialToYMD(Math.floor(v)) : null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return [+m[1], +m[2], +m[3]];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) return [m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]];
  return null;
}
function hora(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') { const f = v - Math.floor(v); return Math.round(f * 1440) % 1440; }
  const s = String(v).trim().toLowerCase();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return null;
  let h = +m[1];
  const pm = /p\.?\s*m/.test(s), am = /a\.?\s*m/.test(s);
  if (pm && h < 12) h += 12;
  if (am && h === 12) h = 0;
  return h * 60 + +m[2];
}
const hhmm = (min) => (min === null || min === undefined ? '' : String(Math.floor(min / 60)).padStart(2, '0') + ':' + String(min % 60).padStart(2, '0'));
function num(v) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return null;
  let s = String(v).trim().replace(/\s/g, '');
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
// Google Sheets a veces convierte "25.4" en la fecha 25-abr: se restaura el decimal
function kg(v) {
  const n = num(v);
  if (n === null) return 0;
  if (n > 20000) { const d = serialToYMD(n); return parseFloat(d[2] + '.' + d[1]); }
  return n;
}
const ym = (f) => f[0] * 100 + f[1];
const dow = (f) => (new Date(Date.UTC(f[0], f[1] - 1, f[2])).getUTCDay() + 6) % 7; // 0 = lunes
const ymd = (f) => f[0] + '-' + String(f[1]).padStart(2, '0') + '-' + String(f[2]).padStart(2, '0');
const minutosEntre = (f1, h1, f2, h2) => (Date.UTC(f2[0], f2[1] - 1, f2[2]) - Date.UTC(f1[0], f1[1] - 1, f1[2])) / 6e4 + (h2 - h1);
function hoyBogota() { const d = new Date(Date.now() - 5 * 36e5); return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]; }

// ---- equivalencias de nombres (tabla del Sheet del portal)
function makeEquiv(rows) {
  const L = rows.slice(1).filter((r) => r[1] && r[2]).map((r) => ({ t: sinTilde(r[0]), c: sinTilde(r[1]), n: String(r[2]).trim() }));
  return (tipo, v) => {
    const s = sinTilde(v);
    if (!s) return '';
    const e = L.find((x) => x.t === tipo && s.indexOf(x.c) >= 0);
    return e ? e.n : String(v).trim();
  };
}
const titulo = (s) => String(s || '').trim().toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());

// ============================================================
// 1) REGISTRO DE CONDUCTORES (App Despachos) — primera hoja del Sheet
//    Columnas por posición, igual que api/action.js de esa app
// ============================================================
async function leerConductores(sheets, eq) {
  const tabs = await titles(sheets, IDS.cond);
  const t = tabs[0];
  const rows = (await readTabs(sheets, IDS.cond, [t]))[t] || [];
  const out = [];
  rows.forEach((r, i) => {
    if (i === 0 && !/^\d+$/.test(String(r[0] || '').trim())) return; // encabezado
    const f = fecha(r[12]);
    if (!f) return;
    if (/prueba/i.test(String(r[2] || '')) || /prueba/i.test(String(r[8] || ''))) return;
    const hi = hora(r[13]), hm = hora(r[14]), hs = hora(r[15]);
    const dif = (a, b) => { if (a === null || b === null) return null; let d = b - a; if (d < 0 && d > -720) return null; if (d < 0) d += 1440; return d; };
    const opr0 = String(r[11] || '').trim(), opr1 = eq('operario', opr0);
    const opr = opr1 ? (opr1 === opr0 ? titulo(opr0) : opr1) : '';
    out.push([ym(f), f[2], dow(f), hhmm(hi),
      /^c/i.test(String(r[1] || '')) ? 'C' : 'D',
      /tecni/i.test(String(r[3] || '')) ? 'T' : 'C',
      eq('cliente', String(r[4] || '').toUpperCase()) || 'SIN CLIENTE',
      /tracto/i.test(String(r[10] || '')) ? 'T' : 'S',
      opr || 'Sin operario',
      dif(hi, hm), dif(hm, hs), dif(hi, hs), String(r[0])]);
  });
  return out;
}

// ============================================================
// 2) REQUISICIÓN A BODEGA DE ROLLOS (App Trazabilidad Rollos)
// ============================================================
async function leerRollos(sheets, eq) {
  const tabs = await titles(sheets, IDS.rol);
  const pares = [['C', findTab(tabs, 'Requisiciones'), findTab(tabs, 'Entregas')], ['T', findTab(tabs, 'Requisiciones_Tecnipapel'), findTab(tabs, 'Entregas_Tecnipapel')]].filter((p) => p[1] && p[2]);
  const data = await readTabs(sheets, IDS.rol, pares.flatMap((p) => [p[1], p[2]]));
  const R = [], E = [], idx = {};
  pares.forEach(([emp, rt, et]) => {
    const rr = data[rt] || [], h = rr[0] || [];
    const c = {
      id: colIdx(h, 'ReqID'), f: colIdx(h, 'FechaCreacion'), h: colIdx(h, 'HoraCreacion'), op: colIdx(h, 'OrdenProduccion'), ref: colIdx(h, 'Referencia'),
      cant: colIdx(h, 'CantidadSolicitada'), maq: colIdx(h, 'MaquinaDestino'), por: colIdx(h, 'CreadoPor'), est: colIdx(h, 'Estado'), dest: colIdx(h, 'DestinoProducto'),
      fc: colIdx(h, 'FechaCierre'), hc: colIdx(h, 'HoraCierre')
    };
    rr.slice(1).forEach((r) => {
      const id = String(cell(r, c.id)).trim(), f = fecha(cell(r, c.f));
      if (!id || !f || /prueba/i.test(String(cell(r, c.op)))) return;
      const hh = hora(cell(r, c.h)), fc = fecha(cell(r, c.fc)), hc = hora(cell(r, c.hc));
      const ciclo = fc && hc !== null && hh !== null ? Math.round(minutosEntre(f, hh, fc, hc) / 6) / 10 : null;
      idx[emp + ':' + id] = R.length;
      R.push([ym(f), f[2], emp, id, Math.round(num(cell(r, c.cant)) || 0), String(cell(r, c.maq)), titulo(cell(r, c.por)), String(cell(r, c.est)), String(cell(r, c.dest)) || 'Sin destino', ciclo, hhmm(hh), null, String(cell(r, c.ref)), emp + ':' + id]);
    });
    const ee = data[et] || [], g = ee[0] || [];
    const d = {
      id: colIdx(g, 'ReqID'), ro: colIdx(g, 'NumeroRollo'), pe: colIdx(g, 'Peso'), cli: colIdx(g, 'Cliente'), maq: colIdx(g, 'MaquinaReal'), cam: colIdx(g, 'HuboCambioMaquina'),
      tipo: colIdx(g, 'Tipo'), met: colIdx(g, 'MetodoCaptura'), opr: colIdx(g, 'OperarioDespacha'), f: colIdx(g, 'FechaEntrega'), h: colIdx(g, 'HoraEntrega'), est: colIdx(g, 'Estado')
    };
    ee.slice(1).forEach((r) => {
      const k = emp + ':' + String(cell(r, d.id)).trim();
      if (!(k in idx)) return;
      const f = fecha(cell(r, d.f));
      if (!f) return;
      const hh = hora(cell(r, d.h)), rollo = String(cell(r, d.ro));
      E.push([ym(f), f[2], emp, idx[k], rollo, Math.round(num(cell(r, d.pe)) || 0), String(cell(r, d.cli)).trim() || 'Sin cliente', String(cell(r, d.maq)),
        /true|si|sí/i.test(String(cell(r, d.cam))) ? 1 : 0, String(cell(r, d.tipo)) || 'Normal', String(cell(r, d.met)), eq('operario', cell(r, d.opr)) || 'Sin operario',
        String(cell(r, d.est)), hhmm(hh), dow(f), k + ':' + rollo + ':' + ymd(f) + ':' + hhmm(hh)]);
    });
  });
  return { R, E };
}

// ============================================================
// 3) CHECKLIST MONTACARGAS (FOR-SST-01)
// ============================================================
async function leerMontacargas(sheets, eq) {
  const tabs = await titles(sheets, IDS.mon);
  const tI = findTab(tabs, 'INSPECCIONES'), tM = findTab(tabs, 'MANTENIMIENTOS'), tF = findTab(tabs, 'FLOTA'), tT = findTab(tabs, 'ITEMS');
  const data = await readTabs(sheets, IDS.mon, [tI, tM, tF, tT].filter(Boolean));
  // FLOTA
  const fr = data[tF] || [], fh = fr[0] || [];
  const fc = { id: colIdx(fh, 'ID'), emp: colIdx(fh, 'Empresa'), tipo: colIdx(fh, 'Tipo'), desc: colIdx(fh, 'Descripcion'), est: colIdx(fh, 'Estado'), hor: colIdx(fh, 'Horometro_actual'), ult: colIdx(fh, 'Horometro_ultimo_mant'), prox: colIdx(fh, 'Prox_mant_h'), soat: colIdx(fh, 'Vencimiento_SOAT'), ext: colIdx(fh, 'Vencimiento_extintor') };
  const F = [], emp = {};
  fr.slice(1).forEach((r) => {
    const id = String(cell(r, fc.id)).trim();
    if (!id) return;
    const e = /tecni/i.test(String(cell(r, fc.emp))) ? 'T' : 'C';
    emp[id] = e;
    const s = fecha(cell(r, fc.soat)), x = fecha(cell(r, fc.ext));
    const nn = (v) => { const n = typeof v === 'number' && v > 20000 ? null : num(v); return n; };
    F.push([id, e, String(cell(r, fc.tipo)), String(cell(r, fc.desc)), String(cell(r, fc.est)) || 'Operativo', nn(cell(r, fc.hor)), nn(cell(r, fc.ult)), nn(cell(r, fc.prox)), s ? ymd(s) : '', x ? ymd(x) : '']);
  });
  // ITEMS (código → nombre)
  const names = {};
  (data[tT] || []).slice(1).forEach((r) => { if (r[0]) names[sinTilde(r[0])] = [String(r[1] || r[0]), String(r[2] || '')]; });
  // INSPECCIONES
  const ir = data[tI] || [], ih = ir[0] || [];
  const ic = { f: colIdx(ih, 'Fecha'), h: colIdx(ih, 'Hora'), op: colIdx(ih, 'Operario', 'Nombre'), maq: colIdx(ih, 'Maquina', 'Máquina'), tipo: colIdx(ih, 'Tipo Equipo', 'Tipo'), hor: colIdx(ih, 'Horometro', 'Horómetro'), obs: colIdx(ih, 'Observaciones') };
  let i0 = ic.hor >= 0 ? ic.hor + 1 : 6;
  let i1 = ih.findIndex((x, j) => j >= i0 && /items en no|total no|items_no|total_no/i.test(sinTilde(x)));
  if (i1 < 0) i1 = Math.min(i0 + 29, ih.length);
  const IT = [];
  for (let j = i0; j < i1; j++) {
    const raw = String(ih[j] || ('Ítem ' + (j - i0 + 1))).replace(/^item\s+/i, '').trim();
    const nm = names[sinTilde(raw)];
    IT.push([raw, nm ? nm[0] : raw, nm ? nm[1] : '']);
  }
  const lastHor = {};
  const I = [];
  ir.slice(1).forEach((r) => {
    const f = fecha(cell(r, ic.f)), maq = String(cell(r, ic.maq)).trim();
    if (!f || !maq) return;
    const no = [];
    let ev = 0;
    for (let j = i0; j < i1; j++) { const v = sinTilde(cell(r, j)); if (v === 's' || v === 'si' || v === 'n' || v === 'no') { ev++; if (v === 'n' || v === 'no') no.push(j - i0); } }
    const hor = num(cell(r, ic.hor)), hh = hora(cell(r, ic.h));
    if (hor !== null && hor < 100000) lastHor[maq] = Math.max(lastHor[maq] || 0, hor);
    const op = eq('operario', cell(r, ic.op)) || 'Sin operario';
    I.push([ym(f), f[2], dow(f), hhmm(hh), op, maq, emp[maq] || 'C', /gato/i.test(String(cell(r, ic.tipo))) ? 'G' : 'M', hor, no, ev, String(cell(r, ic.obs)).trim().slice(0, 160), ymd(f) + ' ' + hhmm(hh) + ' ' + maq + ' ' + op]);
  });
  // El horómetro actual es el mayor entre FLOTA y la última inspección
  F.forEach((x) => { if (lastHor[x[0]] && (x[5] === null || lastHor[x[0]] > x[5])) x[5] = lastHor[x[0]]; });
  // MANTENIMIENTOS
  const mr = data[tM] || [], mh = mr[0] || [];
  const mc = { eq: colIdx(mh, 'Equipo_ID', 'Equipo'), f: colIdx(mh, 'Fecha'), hor: colIdx(mh, 'Horometro_h', 'Horometro'), tipo: colIdx(mh, 'Tipo'), desc: colIdx(mh, 'Descripcion'), prov: colIdx(mh, 'Proveedor_Taller', 'Proveedor'), costo: colIdx(mh, 'Costo') };
  const M = [];
  mr.slice(1).forEach((r) => {
    const f = fecha(cell(r, mc.f)), e = String(cell(r, mc.eq)).trim();
    if (!f || !e) return;
    const tipo = String(cell(r, mc.tipo)).trim(), costo = num(cell(r, mc.costo)) || 0;
    M.push([ym(f), f[2], e, tipo, costo, String(cell(r, mc.desc)).slice(0, 120), String(cell(r, mc.prov)), String(cell(r, mc.hor)), e + ' ' + ymd(f) + ' ' + tipo + ' ' + costo]);
  });
  return { I, M, F, IT };
}

// ============================================================
// 4) CONTROL DE IMPRESOS (solo el módulo de impresos)
// ============================================================
async function leerImpresos(sheets, eq) {
  const tabs = await titles(sheets, IDS.imp);
  const tR = findTab(tabs, 'Impresos_Requisiciones'), tG = findTab(tabs, 'Impresos_Ingresos'), tV = findTab(tabs, 'Impresos_Inventario'), tL = findTab(tabs, 'Impresos_CortesGenerales'), tC = findTab(tabs, 'Impresos_CorteGeneralItems');
  const data = await readTabs(sheets, IDS.imp, [tR, tG, tV, tL, tC].filter(Boolean));
  const tab = (t, must) => { const rows = data[t] || []; const hi = headerRow(rows, must); return { h: rows[hi] || [], rows: rows.slice(hi + 1) }; };
  // Requisiciones
  const q = tab(tR, ['ID', 'Código']);
  const c = { id: colIdx(q.h, 'ID'), f: colIdx(q.h, 'Fecha'), cod: colIdx(q.h, 'Código'), ref: colIdx(q.h, 'Referencia'), sol: colIdx(q.h, 'Solicitó'), cs: colIdx(q.h, 'Cantidad solicitada'), est: colIdx(q.h, 'Estado'), ent: colIdx(q.h, 'Entregó'), rec: colIdx(q.h, 'Entregado a'), ce: colIdx(q.h, 'Cantidad entregada'), cd: colIdx(q.h, 'Cantidad devuelta'), fe: colIdx(q.h, 'Fecha entrega'), h: colIdx(q.h, 'Hora'), he: colIdx(q.h, 'Hora entrega') };
  const R = [];
  q.rows.forEach((r) => {
    const id = String(cell(r, c.id)).trim();
    if (!/^IMP-\d+/i.test(id)) return;
    const f = fecha(cell(r, c.f));
    if (!f) return;
    const fe = fecha(cell(r, c.fe));
    R.push([id, ym(f), f[2], String(cell(r, c.cod)), String(cell(r, c.ref)), titulo(cell(r, c.sol)), kg(cell(r, c.cs)), String(cell(r, c.est)), String(cell(r, c.ent)), titulo(eq('operario', cell(r, c.rec))) || 'Sin receptor', kg(cell(r, c.ce)), kg(cell(r, c.cd)), fe ? ym(fe) : null, fe ? fe[2] : null, hhmm(hora(cell(r, c.h))), hhmm(hora(cell(r, c.he))), 0]);
  });
  // Ingresos
  const g = tab(tG, ['ID', 'Cantidad']);
  const gc = { id: colIdx(g.h, 'ID'), f: colIdx(g.h, 'Fecha'), cod: colIdx(g.h, 'Código'), ref: colIdx(g.h, 'Referencia'), cant: colIdx(g.h, 'Cantidad'), por: colIdx(g.h, 'Ingresó'), obs: colIdx(g.h, 'Observación') };
  const G = [];
  g.rows.forEach((r) => {
    const id = String(cell(r, gc.id)).trim(), f = fecha(cell(r, gc.f));
    if (!id || !f) return;
    G.push([id, ym(f), f[2], String(cell(r, gc.cod)), String(cell(r, gc.ref)), kg(cell(r, gc.cant)), String(cell(r, gc.por)), String(cell(r, gc.obs)).slice(0, 80), 0]);
  });
  // Inventario
  const v = tab(tV, ['Código', 'Referencia']);
  const vc = { cod: colIdx(v.h, 'Código'), ref: colIdx(v.h, 'Referencia'), st: colIdx(v.h, 'Stock actual'), op: colIdx(v.h, 'Óptimo'), cr: colIdx(v.h, 'Crítico') };
  const V = [];
  v.rows.forEach((r) => { const cod = String(cell(r, vc.cod)).trim(); if (!cod) return; V.push([cod, String(cell(r, vc.ref)), kg(cell(r, vc.st)), cell(r, vc.op) === '' ? null : num(cell(r, vc.op)), cell(r, vc.cr) === '' ? null : num(cell(r, vc.cr))]); });
  // Cortes generales (lotes) e ítems contados
  const l = tab(tL, ['ID', 'Estado']);
  const lc = { id: colIdx(l.h, 'ID'), est: colIdx(l.h, 'Estado'), ff: colIdx(l.h, 'Fecha fin'), fi: colIdx(l.h, 'Fecha inicio'), por: colIdx(l.h, 'Iniciado por') };
  const L = [];
  l.rows.forEach((r) => { const id = String(cell(r, lc.id)).trim(); if (!id) return; const ff = fecha(cell(r, lc.ff)), fi = fecha(cell(r, lc.fi)); L.push([id, String(cell(r, lc.est)), ff ? ymd(ff) : '', fi ? ymd(fi) : '', String(cell(r, lc.por))]); });
  const ci = tab(tC, ['Lote', 'Código']);
  const cc = { lote: colIdx(ci.h, 'Lote'), cod: colIdx(ci.h, 'Código'), ref: colIdx(ci.h, 'Referencia'), sis: colIdx(ci.h, 'Cantidad sistema'), con: colIdx(ci.h, 'Cantidad contada'), obs: colIdx(ci.h, 'Observación'), por: colIdx(ci.h, 'Contado por') };
  const C = [];
  ci.rows.forEach((r) => { const lote = String(cell(r, cc.lote)).trim(); if (!lote) return; C.push([lote, String(cell(r, cc.cod)), String(cell(r, cc.ref)), kg(cell(r, cc.sis)), kg(cell(r, cc.con)), String(cell(r, cc.obs)).slice(0, 80), String(cell(r, cc.por)), lote + ':' + String(cell(r, cc.cod))]); });
  return { R, G, V, L, C };
}

// ============================================================
// SHEET DEL PORTAL: se crean las pestañas que falten
// ============================================================
const PORTAL_TABS = {
  Usuarios: [['PIN', 'Nombre', 'Rol', 'Activo'], ['', 'Néstor Solarte', 'Jefe de Logística', 'SI'], ['', 'Alejandro Molina', 'Jefe de Logística', 'SI'], ['', 'Kely Cardona', 'Supervisor de Inventarios', 'SI']],
  Metas: [['Clave', 'Valor', 'Descripción'],
    ['cond_tot', 60, 'Conductores · tiempo total en planta ≤ (min)'], ['cond_esp', 30, 'Conductores · tiempo de espera ≤ (min)'], ['cond_e60', 10, 'Conductores · % vehículos con espera > 1 h ≤'],
    ['rol_cap', 95, 'Rollos · captura automática ≥ (%)'], ['rol_dev', 3, 'Rollos · devoluciones ≤ (%)'],
    ['mon_cov', 95, 'Montacargas · cobertura diaria ≥ (%)'], ['mon_conf', 95, 'Montacargas · índice de conformidad ≥ (%)'], ['mon_prev', 70, 'Montacargas · mantenimiento preventivo ≥ (%)'],
    ['imp_mismo', 95, 'Impresos · entregadas el mismo día ≥ (%)'], ['imp_tol', 95, 'Impresos · exactitud del corte dentro de la tolerancia ≥ (%)'], ['imp_tolpct', 2, 'Impresos · tolerancia del corte ± (%)'], ['imp_abs', 2, 'Impresos · diferencia absoluta / stock ≤ (%)'], ['imp_pdev', '', 'Impresos · % de devolución ≤ (vacío = sin meta)'],
    ['ambar', 10, 'Margen del ámbar: % de la meta dentro del cual se pinta ámbar y no rojo']],
  Cierres: [['Mes', 'Estado', 'Usuario', 'FechaHora', 'Nota']],
  CierreDatos: [['Mes', 'FechaHora', 'Parte', 'Total', 'JSON']],
  Excluidos: [['Clave', 'Activo', 'Usuario', 'FechaHora', 'Motivo'], ['imp:r:IMP-0024', 'SI', 'Sistema', '2026-10-01 00:00', 'Ajuste del ingreso doble de SOCO'], ['imp:g:IMP-ING-0004', 'SI', 'Sistema', '2026-10-01 00:00', 'Ingreso duplicado de SOCO']],
  Equivalencias: [['Tipo', 'Contiene', 'Nombre unificado'],
    ['cliente', 'CARVAJAL', 'CARVAJAL'],
    ['operario', 'UBER', 'Uber Ante'], ['operario', 'GOLU', 'Juan David Golu'], ['operario', 'CARO', 'Julian Caro'], ['operario', 'MIRANDA', 'Alexander Miranda'], ['operario', 'ALEXANDER', 'Alexander Miranda'],
    ['operario', 'IPUS', 'Jhonatan Ipus'], ['operario', 'JHONATAN', 'Jhonatan Ipus'], ['operario', 'CUERVO', 'Henry Cuervo'], ['operario', 'HENRY', 'Henry Cuervo'], ['operario', 'HARRISON', 'Harrison Manzano'], ['operario', 'NESTOR', 'Nestor Solarte']],
  Log: [['FechaHora', 'Usuario', 'Acción', 'Detalle']]
};
let portalReady = false;
async function ensurePortal(sheets) {
  if (portalReady) return;
  const have = await titles(sheets, IDS.portal);
  const falta = Object.keys(PORTAL_TABS).filter((t) => have.indexOf(t) < 0);
  if (falta.length) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: IDS.portal, requestBody: { requests: falta.map((t) => ({ addSheet: { properties: { title: t } } })) } });
    for (const t of falta) {
      await sheets.spreadsheets.values.update({ spreadsheetId: IDS.portal, range: `'${t}'!A1`, valueInputOption: 'RAW', requestBody: { values: PORTAL_TABS[t] } });
    }
  }
  portalReady = true;
}
async function append(sheets, tab, rows) {
  await sheets.spreadsheets.values.append({ spreadsheetId: IDS.portal, range: `'${tab}'!A1`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: rows } });
}
function ahora() { const d = new Date(Date.now() - 5 * 36e5); return d.toISOString().slice(0, 16).replace('T', ' '); }

async function leerPortal(sheets) {
  const d = await readTabs(sheets, IDS.portal, Object.keys(PORTAL_TABS));
  return d;
}
function usuario(P, pin) {
  const p = String(pin || '').trim();
  if (!p) return null;
  const u = (P.Usuarios || []).slice(1).find((r) => String(r[0]).trim() === p && !/^no$/i.test(String(r[3] || 'SI').trim()));
  return u ? { nombre: String(u[1]), rol: String(u[2]) } : null;
}
function estadoPortal(P) {
  const metas = {};
  (P.Metas || []).slice(1).forEach((r) => { if (r[0]) metas[String(r[0]).trim()] = r[1] === '' || r[1] === undefined ? null : num(r[1]); });
  const cierres = {};
  (P.Cierres || []).slice(1).forEach((r) => { if (r[0]) cierres[String(r[0])] = { estado: String(r[1]), usuario: String(r[2]), fh: String(r[3]), nota: String(r[4] || '') }; });
  const hist = (P.Cierres || []).slice(1).filter((r) => r[0]).map((r) => ({ mes: String(r[0]), estado: String(r[1]), usuario: String(r[2]), fh: String(r[3]), nota: String(r[4] || '') }));
  const frozen = {};
  const partes = {};
  (P.CierreDatos || []).slice(1).forEach((r) => {
    const mes = String(r[0]), fh = String(r[1]);
    const c = cierres[mes];
    if (!c || c.estado !== 'Cerrado' || c.fh !== fh) return;
    (partes[mes] = partes[mes] || []).push([Number(r[2]), String(r[4] || '')]);
  });
  Object.keys(partes).forEach((mes) => { try { frozen[mes] = JSON.parse(partes[mes].sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('')); } catch (e) { /* cierre dañado: se ignora */ } });
  const excl = {};
  (P.Excluidos || []).slice(1).forEach((r) => { if (r[0]) excl[String(r[0])] = { activo: !/^no$/i.test(String(r[1]).trim()), usuario: String(r[2]), fh: String(r[3]), motivo: String(r[4] || '') }; });
  return { metas, cierres, hist, frozen, excl };
}

let cache = { t: 0, data: null };
async function leerTodo(sheets, eq) {
  if (cache.data && Date.now() - cache.t < 60000) return cache.data;
  const out = { errores: {} };
  const tareas = [
    ['cond', () => leerConductores(sheets, eq)], ['rol', () => leerRollos(sheets, eq)],
    ['mon', () => leerMontacargas(sheets, eq)], ['imp', () => leerImpresos(sheets, eq)]
  ];
  await Promise.all(tareas.map(async ([k, fn]) => {
    if (!IDS[k]) { out.errores[k] = 'Falta la variable SHEET_ID de este proceso en Vercel.'; return; }
    try { out[k] = await fn(); } catch (e) { out.errores[k] = String((e && e.message) || e); }
  }));
  cache = { t: Date.now(), data: out };
  return out;
}

// ============================================================
// HANDLER
// ============================================================
module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') { res.status(405).json({ ok: false, error: 'Use POST' }); return; }
  if (!ENV.GOOGLE_SERVICE_ACCOUNT_EMAIL || !ENV.GOOGLE_PRIVATE_KEY || !IDS.portal) {
    res.status(500).json({ ok: false, error: 'Faltan variables en Vercel: GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY y SHEET_ID_PORTAL.' });
    return;
  }
  let b = req.body;
  if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
  b = b || {};
  try {
    const sheets = await getSheets();
    await ensurePortal(sheets);
    const P = await leerPortal(sheets);
    const u = usuario(P, b.pin);
    if (!u) { res.status(200).json({ ok: false, error: 'PIN incorrecto' }); return; }

    if (b.action === 'login') { res.status(200).json({ ok: true, user: u }); return; }

    if (b.action === 'datos') {
      const eq = makeEquiv(P.Equivalencias || []);
      const data = await leerTodo(sheets, eq);
      const h = hoyBogota();
      res.status(200).json({ ok: true, user: u, hoy: ymd(h), data, portal: estadoPortal(P) });
      return;
    }

    if (b.action === 'excluir') {
      if (!b.clave) throw new Error('Falta la clave del registro');
      await append(sheets, 'Excluidos', [[String(b.clave), b.activo ? 'SI' : 'NO', u.nombre, ahora(), String(b.motivo || '').slice(0, 200)]]);
      await append(sheets, 'Log', [[ahora(), u.nombre, b.activo ? 'Excluir' : 'Incluir', String(b.clave)]]);
      res.status(200).json({ ok: true, fh: ahora(), usuario: u.nombre });
      return;
    }

    if (b.action === 'metas') {
      const metas = b.metas || {};
      const rows = (P.Metas || []).slice();
      const keys = rows.map((r) => String(r[0] || ''));
      const nuevas = [];
      Object.keys(metas).forEach((k) => { const i = keys.indexOf(k); const v = metas[k] === null || metas[k] === '' ? '' : Number(metas[k]); if (i > 0) rows[i] = [k, v, rows[i][2] || '']; else nuevas.push([k, v, '']); });
      await sheets.spreadsheets.values.update({ spreadsheetId: IDS.portal, range: `'Metas'!A1`, valueInputOption: 'RAW', requestBody: { values: rows.concat(nuevas).map((r) => [r[0], r[1] === undefined ? '' : r[1], r[2] || '']) } });
      await append(sheets, 'Log', [[ahora(), u.nombre, 'Cambiar metas', JSON.stringify(metas).slice(0, 400)]]);
      res.status(200).json({ ok: true });
      return;
    }

    if (b.action === 'cerrar') {
      const mes = String(b.mes || ''), json = String(b.json || '');
      if (!/^\d{6}$/.test(mes) || !json) throw new Error('Datos del cierre incompletos');
      const fh = ahora();
      const partes = [];
      for (let i = 0; i < json.length; i += 45000) partes.push(json.slice(i, i + 45000));
      await append(sheets, 'CierreDatos', partes.map((p, i) => [mes, fh, i, partes.length, p]));
      await append(sheets, 'Cierres', [[mes, 'Cerrado', u.nombre, fh, String(b.nota || '')]]);
      await append(sheets, 'Log', [[fh, u.nombre, 'Cerrar mes', mes]]);
      res.status(200).json({ ok: true, fh, usuario: u.nombre });
      return;
    }

    if (b.action === 'reabrir') {
      const mes = String(b.mes || '');
      if (!/^\d{6}$/.test(mes)) throw new Error('Mes inválido');
      const fh = ahora();
      await append(sheets, 'Cierres', [[mes, 'Reabierto', u.nombre, fh, String(b.nota || '')]]);
      await append(sheets, 'Log', [[fh, u.nombre, 'Reabrir mes', mes + ' · ' + String(b.nota || '')]]);
      res.status(200).json({ ok: true, fh, usuario: u.nombre });
      return;
    }

    res.status(200).json({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    console.error('api/indicadores', err);
    res.status(500).json({ ok: false, error: String((err && err.message) || err) });
  }
};

// Para las pruebas locales
module.exports._t = { fecha, hora, num, kg, makeEquiv, leerConductores, leerRollos, leerMontacargas, leerImpresos, estadoPortal, PORTAL_TABS, IDS };
