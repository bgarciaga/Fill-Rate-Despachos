const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const carpeta = "./excels";
const HOJA = "Fillrate por despacho";

const columnas = [
  "pedido", "codigo", "ean", "descripcion",
  "u_pedidas", "u_anuladas", "u_servidas", "pte_servir",
  "stock", "ubicacion", "bultos_pedidos", "bultos_servidos",
  "proveedor"
];

// Respaldo si no se encuentra un encabezado
const COL_DEFAULT = {
  pedido: 1, codigo: 3, ean: 7, descripcion: 8,
  bultos_pedidos: 10, bultos_servidos: 11,
  u_pedidas: 12, u_anuladas: 13, u_servidas: 14, pte_servir: 15,
  stock: 17, ubicacion: 24, proveedor: -1
};

const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
const num = v => { if (typeof v === "number") return v; const n = parseFloat(String(v ?? "").replace(/,/g, "").trim()); return isNaN(n) ? 0 : n; };
const txt = v => String(v ?? "").replace(/\s+/g, " ").trim();
const pedidoId = v => txt(v).replace(/\.0+$/, "").replace(/^0+/, "");
const pad = n => String(n).padStart(2, "0");

// ---------- FECHAS ----------
function parseFecha(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") {
    if (v < 44000 || v > 50000) return null; // serial Excel 2020–2036
    const p = XLSX.SSF.parse_date_code(v);
    return p ? { fecha: `${p.y}-${pad(p.m)}-${pad(p.d)}`, hora: `${pad(p.H)}:${pad(p.M)}` } : null;
  }
  if (v instanceof Date && !isNaN(v)) {
    return { fecha: `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`, hora: `${pad(v.getHours())}:${pad(v.getMinutes())}` };
  }
  const s = txt(v);
  let m;
  // yyyy-mm-dd
  m = s.match(/(20\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) { const [, y, mo, d, H = "0", M = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` }; }
  // dd/mm/yyyy
  m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) { const [, d, mo, y, H = "0", M = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` }; }
  // yy/mm/dd  (formato del reporte: 26/10/02 = 2026-10-02)
  m = s.match(/(\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) { const [, y, mo, d, H = "0", M = "0"] = m; return { fecha: `20${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` }; }
  return null;
}

function buscarValor(data, palabras, filas = 4) {
  for (let r = 0; r < Math.min(filas, data.length); r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (palabras.some(p => norm(row[c]).includes(p))) {
        for (let k = c + 1; k < row.length; k++) if (row[k] !== "") return row[k];
      }
    }
  }
  return "";
}

// ---------- COLUMNAS POR ENCABEZADO ----------
const ffill = row => { let last = ""; return (row || []).map(v => (txt(v) ? (last = txt(v)) : last)); };

function detectarColumnas(data) {
  // fila que contiene "descripcion" = encabezado principal
  let hr = -1;
  for (let r = 0; r < Math.min(15, data.length); r++) {
    if ((data[r] || []).some(c => norm(c).includes("descripcion"))) { hr = r; break; }
  }
  if (hr < 0) return { inicio: 4, idx: { ...COL_DEFAULT }, labels: [] };

  const ancho = Math.max(...data.slice(hr, hr + 3).map(r => (r || []).length));
  const arriba = ffill(data[hr - 1]);            // grupo arriba (ej. "Unidades", "Bultos")
  const main = data[hr] || [];
  const sub = data[hr + 1] || [];
  // ¿la fila siguiente es sub-encabezado? (texto, sin números)
  const subEsHeader = sub.some(c => /[a-z]/i.test(String(c))) && !sub.some(c => typeof c === "number");
  const mainFill = ffill(main);

  const labels = [];
  for (let c = 0; c < ancho; c++) {
    labels[c] = subEsHeader
      ? norm(`${mainFill[c]} ${txt(sub[c])}`)
      : norm(`${txt(main[c]) ? "" : ""}${arriba[c] && arriba[c] !== txt(main[c]) ? arriba[c] + " " : ""}${txt(main[c])}`);
  }
  const inicio = hr + (subEsHeader ? 2 : 1);

  const usados = new Set();
  const find = (test) => {
    for (let c = 0; c < labels.length; c++) {
      if (usados.has(c) || !labels[c]) continue;
      if (test(labels[c])) { usados.add(c); return c; }
    }
    return -1;
  };
  const isB = l => /bult|bto|caja/.test(l);

  const idx = {};
  idx.descripcion     = find(l => /descripcion/.test(l));
  idx.ean             = find(l => /\bean\b|barra/.test(l));
  idx.proveedor       = find(l => /proveedor/.test(l));
  idx.ubicacion       = find(l => /ubic/.test(l));
  idx.stock           = find(l => /stock|existencia/.test(l));
  idx.bultos_pedidos  = find(l => isB(l) && /ped/.test(l));
  idx.bultos_servidos = find(l => isB(l) && /serv/.test(l));
  idx.u_anuladas      = find(l => !isB(l) && /anul/.test(l));
  idx.pte_servir      = find(l => !isB(l) && /pte|pend/.test(l));
  idx.u_servidas      = find(l => !isB(l) && /serv/.test(l));
  idx.u_pedidas       = find(l => !isB(l) && /pedid[ao]s|u\.? ?ped|unid.*ped|cant.*ped/.test(l));
  idx.pedido          = find(l => /pedido/.test(l));
  idx.codigo          = find(l => /codigo|articulo|sku/.test(l));

  for (const k of Object.keys(COL_DEFAULT)) if (idx[k] == null || idx[k] < 0) idx[k] = COL_DEFAULT[k];
  return { inicio, idx, labels };
}

// ---------- PROCESO ----------
console.log("Carpeta:", path.resolve(carpeta));
if (!fs.existsSync(carpeta)) { console.log("❌ La carpeta no existe"); process.exit(1); }

const archivos = fs.readdirSync(carpeta)
  .filter(f => /\.xlsx?$/i.test(f) && !f.startsWith("~$"))
  .sort();
console.log("Archivos Excel:", archivos);

const despachos = [];
const ids = new Set();

for (const archivo of archivos) {
  const ruta = path.join(carpeta, archivo);
  console.log(`\n=== ${archivo} ===`);

  let workbook;
  try { workbook = XLSX.readFile(ruta); }
  catch (e) { console.log("❌ No se pudo leer:", e.message); continue; }

  const nombreHoja = workbook.SheetNames.find(n => norm(n) === norm(HOJA))
    || workbook.SheetNames.find(n => norm(n).includes("fillrate") || norm(n).includes("fill rate"))
    || workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHoja], { header: 1, defval: "", raw: true });
  data.slice(0, 7).forEach((r, i) => console.log(`  fila ${i}:`, JSON.stringify(r)));

  const tienda = txt(data?.[1]?.[0]).toUpperCase() || "SIN TIENDA";
  const numero = archivo.match(/(\d+)\.xlsx?$/i)?.[1] || "";

  let fl = parseFecha(buscarValor(data, ["lanzado", "lanzamiento"])) || parseFecha(data?.[2]?.[6]);
  if (!fl) {
    outer: for (let r = 0; r < Math.min(4, data.length); r++)
      for (const c of data[r] || []) { fl = parseFecha(c); if (fl) break outer; }
  }
  if (!fl) {
    console.log("⚠ Sin fecha, uso fecha del archivo");
    const m = fs.statSync(ruta).mtime;
    fl = { fecha: `${m.getFullYear()}-${pad(m.getMonth() + 1)}-${pad(m.getDate())}`, hora: `${pad(m.getHours())}:${pad(m.getMinutes())}` };
  }
  const { fecha, hora } = fl;

  const fp = parseFecha(buscarValor(data, ["fin preparacion", "fin de preparacion", "finalizado"]));
  const finPreparacion = fp ? `${fp.fecha}T${fp.hora}` : "";

  const { inicio, idx, labels } = detectarColumnas(data);
  console.log("Encabezados:", JSON.stringify(labels.map((l, i) => `${i}:${l}`).filter(x => !x.endsWith(":"))));
  console.log("Columnas usadas:", JSON.stringify(idx), "· datos desde fila", inicio);

  const lineas = [];
  for (let i = inicio; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[idx.pedido]) continue;
    const pedido = pedidoId(row[idx.pedido]);
    if (!pedido || !/\d/.test(pedido) || norm(row[idx.pedido]).includes("total")) continue;
    lineas.push([
      pedido,
      txt(row[idx.codigo]),
      txt(row[idx.ean]).replace(/\.0+$/, ""),
      txt(row[idx.descripcion]),
      num(row[idx.u_pedidas]),
      num(row[idx.u_anuladas]),
      num(row[idx.u_servidas]),
      num(row[idx.pte_servir]),
      num(row[idx.stock]),
      txt(row[idx.ubicacion]),
      num(row[idx.bultos_pedidos]),
      num(row[idx.bultos_servidos]),
      idx.proveedor >= 0 ? txt(row[idx.proveedor]).toUpperCase() : ""
    ]);
  }

  const totPed = lineas.reduce((a, l) => a + l[4], 0);
  console.log(`Tienda: ${tienda} · Fecha: ${fecha} ${hora} · Líneas: ${lineas.length} · U. pedidas: ${totPed}`);
  if (lineas.length) console.log("  1ª línea:", JSON.stringify(lineas[0]));
  if (!totPed) console.log("⚠ U. pedidas = 0, revisar columnas");
  if (!lineas.length) { console.log("❌ Sin líneas, se omite"); continue; }

  const id = `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;
  if (ids.has(id)) { console.log("❌ Duplicado, se omite"); continue; }
  ids.add(id);

  despachos.push({ id, numero, archivo, tienda, fecha, hora, finPreparacion, lineas });
}

despachos.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));

fs.writeFileSync("data.json", JSON.stringify({ generado: new Date().toISOString(), columnas, despachos }));
console.log(`\nJSON generado con ${despachos.length} despachos`);
