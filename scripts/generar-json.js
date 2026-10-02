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

const MAPA = {
  pedido:          [1,  ["pedido"]],
  codigo:          [3,  ["codigo", "articulo", "sku"]],
  ean:             [7,  ["ean", "barra"]],
  descripcion:     [8,  ["descripcion"]],
  bultos_pedidos:  [10, ["bultos pedidos", "bul. pedidos", "bultos ped"]],
  bultos_servidos: [11, ["bultos servidos", "bul. servidos", "bultos serv"]],
  u_pedidas:       [12, ["u. pedidas", "unidades pedidas", "u pedidas"]],
  u_anuladas:      [13, ["u. anuladas", "unidades anuladas", "u anuladas"]],
  u_servidas:      [14, ["u. servidas", "unidades servidas", "u servidas"]],
  pte_servir:      [15, ["pte", "pendiente"]],
  stock:           [17, ["stock"]],
  ubicacion:       [24, ["ubicacion"]],
  proveedor:       [-1, ["proveedor"]]
};

const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
const num = v => { if (typeof v === "number") return v; const n = parseFloat(String(v ?? "").replace(/,/g, "").trim()); return isNaN(n) ? 0 : n; };
const txt = v => String(v ?? "").replace(/\s+/g, " ").trim();
const pedidoId = v => txt(v).replace(/\.0+$/, "").replace(/^0+/, "");
const pad = n => String(n).padStart(2, "0");

function parseFecha(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") {
    if (v < 30000 || v > 70000) return null;
    const p = XLSX.SSF.parse_date_code(v);
    return p ? { fecha: `${p.y}-${pad(p.m)}-${pad(p.d)}`, hora: `${pad(p.H)}:${pad(p.M)}` } : null;
  }
  if (v instanceof Date && !isNaN(v)) {
    return { fecha: `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`, hora: `${pad(v.getHours())}:${pad(v.getMinutes())}` };
  }
  const s = txt(v);
  let m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) { const [, y, mo, d, H = "0", M = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` }; }
  m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) { let [, d, mo, y, H = "0", M = "0"] = m; if (y.length === 2) y = "20" + y; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` }; }
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

function detectarColumnas(data) {
  let filaHeader = -1; const idx = {};
  for (let r = 0; r < Math.min(10, data.length); r++) {
    const row = (data[r] || []).map(norm);
    if (row.some(c => c.includes("pedido")) && row.some(c => c.includes("descripcion"))) {
      filaHeader = r;
      for (const [k, [, keys]] of Object.entries(MAPA)) {
        const i = row.findIndex(c => keys.some(key => c.includes(key)));
        if (i >= 0) idx[k] = i;
      }
      break;
    }
  }
  for (const [k, [def]] of Object.entries(MAPA)) if (idx[k] == null) idx[k] = def;
  return { filaHeader, idx };
}

// ---------- DIAGNÓSTICO ----------
console.log("Carpeta:", path.resolve(carpeta));
if (!fs.existsSync(carpeta)) { console.log("❌ La carpeta no existe"); process.exit(1); }
console.log("Contenido:", fs.readdirSync(carpeta));

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

  console.log("Hojas:", workbook.SheetNames);
  const nombreHoja = workbook.SheetNames.find(n => norm(n) === norm(HOJA))
    || workbook.SheetNames.find(n => norm(n).includes("fillrate") || norm(n).includes("fill rate"))
    || workbook.SheetNames[0];
  console.log("Usando hoja:", nombreHoja);
  const hoja = workbook.Sheets[nombreHoja];

  const data = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "", raw: true });
  console.log("Filas:", data.length);
  data.slice(0, 7).forEach((r, i) => console.log(`  fila ${i}:`, JSON.stringify(r)));

  const tienda = txt(data?.[1]?.[0]).toUpperCase() || "SIN TIENDA";
  const numero = archivo.match(/(\d+)\.xlsx?$/i)?.[1] || "";

  let fl = parseFecha(data?.[2]?.[6]) || parseFecha(buscarValor(data, ["lanzado", "lanzamiento"]));
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

  const { filaHeader, idx } = detectarColumnas(data);
  const inicio = filaHeader >= 0 ? filaHeader + 1 : 4;
  console.log("Fila encabezado:", filaHeader, "· columnas:", JSON.stringify(idx));

  const lineas = [];
  for (let i = inicio; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const pedido = pedidoId(row[idx.pedido]);
    if (!pedido || norm(row[idx.pedido]).includes("total")) continue;
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

  console.log(`Tienda: ${tienda} · Fecha: ${fecha} ${hora} · Líneas: ${lineas.length}`);
  if (!lineas.length) { console.log("❌ Sin líneas, se omite"); continue; }

  const id = `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;
  if (ids.has(id)) { console.log("❌ Duplicado, se omite"); continue; }
  ids.add(id);

  despachos.push({ id, numero, archivo, tienda, fecha, hora, finPreparacion, lineas });
  console.log("✔ Agregado");
}

despachos.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));

fs.writeFileSync("data.json", JSON.stringify({ generado: new Date().toISOString(), columnas, despachos }));
console.log(`\nJSON generado con ${despachos.length} despachos`);
