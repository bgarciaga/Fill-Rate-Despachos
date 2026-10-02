const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const carpeta = "./excels";

const columnas = [
  "pedido", "codigo", "ean", "descripcion",
  "u_pedidas", "u_anuladas", "u_servidas", "pte_servir",
  "stock", "ubicacion", "bultos_pedidos", "bultos_servidos",
  "proveedor"
];

// Lee una celda por dirección ("F2", "T5"...)
const cel = (hoja, ref) => {
  const c = hoja[ref];
  return c ? (c.w ?? c.v ?? "") : "";
};
const txt = v => String(v ?? "").replace(/\s+/g, " ").trim();
const num = v => {
  if (typeof v === "number") return v;
  const n = parseFloat(String(v ?? "").replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
};
const pad = n => String(n).padStart(2, "0");

// "1 Cjs, 0 Pcks, 0 Uni" → 1  (bultos = cajas)
const cajas = v => {
  const m = String(v ?? "").match(/(\d+)\s*Cj/i);
  return m ? Number(m[1]) : num(v);
};

// "2026/10/02 09:44:46" → { fecha:"2026-10-02", hora:"09:44", seg:"46" }
function parseFecha(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") {
    const p = XLSX.SSF.parse_date_code(v);
    return p ? { fecha: `${p.y}-${pad(p.m)}-${pad(p.d)}`, hora: `${pad(p.H)}:${pad(p.M)}`, seg: pad(Math.floor(p.S)) } : null;
  }
  const s = txt(v);
  let m = s.match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) { const [, y, mo, d, H = "0", M = "0", S = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}`, seg: pad(S) }; }
  m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) { const [, d, mo, y, H = "0", M = "0", S = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}`, seg: pad(S) }; }
  return null;
}

const archivos = fs.readdirSync(carpeta)
  .filter(f => /\.xlsx$/i.test(f) && !f.startsWith("~$"))
  .sort();

const despachos = [];
const ids = new Set();

for (const archivo of archivos) {
  const ruta = path.join(carpeta, archivo);

  let wb;
  try { wb = XLSX.readFile(ruta); }
  catch (e) { console.log(`❌ ${archivo}: ${e.message}`); continue; }

  const hoja = wb.Sheets["Fillrate por despacho"] || wb.Sheets[wb.SheetNames[0]];
  if (!hoja || !hoja["!ref"]) { console.log(`❌ ${archivo}: hoja vacía`); continue; }

  // ---- Encabezado ----
  const tienda = txt(cel(hoja, "A2")).toUpperCase();
  const numero =
    txt(cel(hoja, "A1")).match(/(\d+)\s*$/)?.[1] ||
    archivo.match(/(\d+)\.xlsx$/i)?.[1] || "";

  const fl = parseFecha(cel(hoja, "F2"));
  if (!fl) { console.log(`❌ ${archivo}: sin fecha lanzado en F2 (${cel(hoja, "F2")})`); continue; }
  const { fecha, hora } = fl;
  const lanzado = `${fecha.replace(/-/g, "/")} ${hora}:${fl.seg}`;

  const fp = parseFecha(cel(hoja, "N2"));
  const finPreparacion = fp ? `${fp.fecha} ${fp.hora}` : "";

  // ---- Líneas (desde fila 5) ----
  const ultima = XLSX.utils.decode_range(hoja["!ref"]).e.r + 1;
  const lineas = [];

  for (let r = 5; r <= ultima; r++) {
    const pedido = txt(cel(hoja, `B${r}`)).replace(/^0+/, "");
    if (!pedido || !/^\d+$/.test(pedido)) continue;

    lineas.push([
      pedido,                                   // pedido
      txt(cel(hoja, `D${r}`)),                  // codigo
      txt(cel(hoja, `J${r}`)),                  // ean
      txt(cel(hoja, `O${r}`)),                  // descripcion
      num(cel(hoja, `T${r}`)),                  // u_pedidas
      num(cel(hoja, `U${r}`)),                  // u_anuladas
      num(cel(hoja, `V${r}`)),                  // u_servidas
      num(cel(hoja, `W${r}`)),                  // pte_servir
      num(cel(hoja, `Y${r}`)),                  // stock
      txt(cel(hoja, `AB${r}`)) || txt(cel(hoja, `AA${r}`)), // ubicacion
      cajas(cel(hoja, `P${r}`)),                // bultos_pedidos
      cajas(cel(hoja, `S${r}`)),                // bultos_servidos
      txt(cel(hoja, `E${r}`)).toUpperCase()     // proveedor
    ]);
  }

  if (!lineas.length) { console.log(`❌ ${archivo}: sin líneas`); continue; }

  const id = `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;
  if (ids.has(id)) { console.log(`⚠ ${archivo}: duplicado, se omite`); continue; }
  ids.add(id);

  despachos.push({ id, numero, archivo, tienda, fecha, hora, lanzado, finPreparacion, lineas });

  const up = lineas.reduce((a, l) => a + l[4], 0);
  const us = lineas.reduce((a, l) => a + l[6], 0);
  console.log(`✔ ${archivo} · ${tienda} · ${lanzado} · ${lineas.length} líneas · ${us}/${up} u.`);
}

despachos.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));

fs.writeFileSync("data.json", JSON.stringify({
  generado: new Date().toISOString(),
  columnas,
  despachos
}));

console.log(`\nJSON generado con ${despachos.length} despachos`);
