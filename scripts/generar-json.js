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

// índice por defecto (tu layout actual) + palabras para detectar el encabezado
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

const norm = s => String(s ?? "")
  .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
  .replace(/\s+/g, " ").trim().toLowerCase();

const num = v => {
  if (typeof v === "number") return v;
  const s = String(v ?? "").replace(/,/g, "").trim();
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};

const txt = v => String(v ?? "").replace(/\s+/g, " ").trim();

const pedidoId = v => txt(v).replace(/\.0+$/, "").replace(/^0+/, "");

const pad = n => String(n).padStart(2, "0");

// Devuelve { fecha: "YYYY-MM-DD", hora: "HH:MM" } en hora LOCAL (sin UTC)
function parseFecha(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") {
    const p = XLSX.SSF.parse_date_code(v);
    if (!p) return null;
    return { fecha: `${p.y}-${pad(p.m)}-${pad(p.d)}`, hora: `${pad(p.H)}:${pad(p.M)}` };
  }
  if (v instanceof Date && !isNaN(v)) {
    return { fecha: `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`,
             hora: `${pad(v.getHours())}:${pad(v.getMinutes())}` };
  }
  const s = txt(v);
  // dd/mm/yyyy hh:mm(:ss) o dd-mm-yyyy
  let m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) {
    let [, d, mo, y, H = "0", M = "0"] = m;
    if (y.length === 2) y = "20" + y;
    return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` };
  }
  // yyyy-mm-dd hh:mm
  m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]+(\d{1,2}):(\d{2}))?/);
  if (m) {
    const [, y, mo, d, H = "0", M = "0"] = m;
    return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}` };
  }
  return null;
}

// Busca una etiqueta en las primeras filas y devuelve el valor de la celda siguiente con dato
function buscarValor(data, palabras, filas = 4) {
  for (let r = 0; r < Math.min(filas, data.length); r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      const n = norm(row[c]);
      if (palabras.some(p => n.includes(p))) {
        for (let k = c + 1; k < row.length; k++) if (row[k] !== "") return row[k];
      }
    }
  }
  return "";
}

function detectarColumnas(data) {
  let filaHeader = -1, idx = {};
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

const despachos = [];
const ids = new Set();

const archivos = fs.readdirSync(carpeta)
  .filter(f => f.toLowerCase().endsWith(".xlsx") && !f.startsWith("~$"))
  .sort();

for (const archivo of archivos) {
  const ruta = path.join(carpeta, archivo);
  let workbook;
  try { workbook = XLSX.readFile(ruta, { cellDates: false }); }
  catch (e) { console.log(`No se pudo leer ${archivo}: ${e.message}`); continue; }

  const nombreHoja = workbook.SheetNames.find(n => norm(n) === norm(HOJA));
  const hoja = nombreHoja && workbook.Sheets[nombreHoja];
  if (!hoja) { console.log(`No existe la hoja en ${archivo}`); continue; }

  const data = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "", raw: true });

  const tienda = txt(data?.[1]?.[0]).toUpperCase();
  const numero = archivo.match(/(\d+)\.xlsx$/i)?.[1] || "";

  const fl = parseFecha(buscarValor(data, ["lanzado", "lanzamiento"]) || data?.[2]?.[6]);
  if (!fl) { console.log(`Sin fecha válida en ${archivo}, se omite`); continue; }
  const { fecha, hora } = fl;

  const fp = parseFecha(buscarValor(data, ["fin preparacion", "fin de preparacion", "finalizado"]));
  const finPreparacion = fp ? `${fp.fecha}T${fp.hora}` : "";

  const { filaHeader, idx } = detectarColumnas(data);
  const inicio = filaHeader >= 0 ? filaHeader + 1 : 4;

  const lineas = [];
  for (let i = inicio; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const pedido = pedidoId(row[idx.pedido]);
    if (!pedido || !/\d/.test(pedido)) continue;          // salta totales / filas vacías
    if (norm(row[idx.pedido]).includes("total")) continue;

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

  if (!lineas.length) { console.log(`Sin líneas en ${archivo}`); continue; }

  let id = `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;
  if (ids.has(id)) { console.log(`Duplicado ${id} (${archivo}), se omite`); continue; }
  ids.add(id);

  despachos.push({ id, numero, archivo, tienda, fecha, hora, finPreparacion, lineas });
  console.log(`✔ ${archivo} → ${tienda} ${fecha} ${hora} · ${lineas.length} líneas · cols:`, idx);
}

despachos.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));

fs.writeFileSync("data.json", JSON.stringify({
  generado: new Date().toISOString(),
  columnas,
  despachos
}));

console.log(`JSON generado con ${despachos.length} despachos`);
