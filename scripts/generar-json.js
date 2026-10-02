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

// Nombre exacto del encabezado en el Excel (sin acentos, mayúsculas)
const HEADERS = {
  pedido:          ["PEDIDO"],
  codigo:          ["CODE ERP", "CODIGO"],
  proveedor:       ["PROVEEDOR"],
  ean:             ["EAN"],
  descripcion:     ["DESCRIPCION"],
  bultos_pedidos:  ["FORMATOS PEDIDOS"],
  bultos_servidos: ["FORMATOS SERVIDOS"],
  u_pedidas:       ["U_PEDIDAS"],
  u_anuladas:      ["U_ANULADAS"],
  u_servidas:      ["U_SERVIDAS"],
  pte_servir:      ["PTE. SERVIR", "PTE SERVIR"],
  stock:           ["STOCK"],
  ubicacion:       ["UBICACIONES"],
  ubiactual:       ["UBIACTUAL"]
};

const up = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
const txt = v => String(v ?? "").replace(/\s+/g, " ").trim();
const num = v => { if (typeof v === "number") return v; const n = parseFloat(String(v ?? "").replace(/,/g, "").trim()); return isNaN(n) ? 0 : n; };
const pad = n => String(n).padStart(2, "0");
const cajas = v => { const m = String(v ?? "").match(/(\d+)\s*Cj/i); return m ? Number(m[1]) : num(v); };

function parseFecha(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") {
    if (v < 44000 || v > 50000) return null;
    const p = XLSX.SSF.parse_date_code(v);
    return p ? { fecha: `${p.y}-${pad(p.m)}-${pad(p.d)}`, hora: `${pad(p.H)}:${pad(p.M)}`, seg: pad(Math.floor(p.S)) } : null;
  }
  const s = txt(v);
  let m = s.match(/(20\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) { const [, y, mo, d, H = "0", M = "0", S = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}`, seg: pad(S) }; }
  m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) { const [, d, mo, y, H = "0", M = "0", S = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}`, seg: pad(S) }; }
  return null;
}

const archivos = fs.readdirSync(carpeta).filter(f => /\.xlsx$/i.test(f) && !f.startsWith("~$")).sort();
const despachos = [];
const ids = new Set();

for (const archivo of archivos) {
  let wb;
  try { wb = XLSX.readFile(path.join(carpeta, archivo)); }
  catch (e) { console.log(`❌ ${archivo}: ${e.message}`); continue; }

  const hoja = wb.Sheets["Fillrate por despacho"] || wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "", raw: true, blankrows: true });

  // ---- Fila de encabezados: la que tenga U_PEDIDAS ----
  const hr = data.findIndex((r, i) => i < 20 && (r || []).some(c => up(c) === "U_PEDIDAS"));
  if (hr < 0) { console.log(`❌ ${archivo}: no encontré la fila de encabezados`); continue; }
  const H = (data[hr] || []).map(up);

  const col = {};
  for (const [k, nombres] of Object.entries(HEADERS)) col[k] = H.findIndex(h => nombres.includes(h));
  const faltan = Object.keys(HEADERS).filter(k => col[k] < 0 && k !== "ubiactual");
  if (faltan.length) console.log(`⚠ ${archivo}: faltan columnas ${faltan.join(", ")}`);

  // ---- Encabezado del despacho (filas antes de los encabezados) ----
  const arriba = data.slice(0, hr);
  const titulo = txt((arriba[0] || []).find(c => /despacho/i.test(String(c))) || "");
  const numero = titulo.match(/(\d+)\s*$/)?.[1] || archivo.match(/(\d+)\.xlsx$/i)?.[1] || "";
  const tienda = up((arriba[1] || []).find(c => txt(c) && !parseFecha(c)) || "");

  // fechas: buscar etiquetas "FECHA LANZADO" / "FECHA FIN" y tomar la fecha más cercana debajo
  const fechas = [];
  arriba.forEach((r, ri) => (r || []).forEach((c, ci) => { const f = parseFecha(c); if (f) fechas.push({ ...f, ri, ci }); }));
  const buscarEtiqueta = re => { for (let ri = 0; ri < arriba.length; ri++) { const ci = (arriba[ri] || []).findIndex(c => re.test(up(c))); if (ci >= 0) return { ri, ci }; } return null; };
  const cercana = et => et && fechas.filter(f => f.ri > et.ri).sort((a, b) => Math.abs(a.ci - et.ci) - Math.abs(b.ci - et.ci))[0];

  const fl = cercana(buscarEtiqueta(/LANZADO/)) || fechas[0];
  const fpTmp = cercana(buscarEtiqueta(/FIN/));
  const fp = fpTmp && fpTmp !== fl ? fpTmp : null;
  if (!fl) { console.log(`❌ ${archivo}: sin fecha lanzado`); continue; }

  const { fecha, hora } = fl;
  const lanzado = `${fecha.replace(/-/g, "/")} ${hora}:${fl.seg}`;
  const finPreparacion = fp ? `${fp.fecha} ${fp.hora}` : "";

  // ---- Líneas ----
  const g = (row, k) => (col[k] >= 0 ? row[col[k]] : "");
  const lineas = [];
  for (let i = hr + 1; i < data.length; i++) {
    const row = data[i] || [];
    const pedido = txt(g(row, "pedido")).replace(/^0+/, "");
    if (!pedido || !/^\d+$/.test(pedido)) continue;
    lineas.push([
      pedido,
      txt(g(row, "codigo")),
      txt(g(row, "ean")),
      txt(g(row, "descripcion")),
      num(g(row, "u_pedidas")),
      num(g(row, "u_anuladas")),
      num(g(row, "u_servidas")),
      num(g(row, "pte_servir")),
      num(g(row, "stock")),
      txt(g(row, "ubicacion")) || txt(g(row, "ubiactual")),
      cajas(g(row, "bultos_pedidos")),
      cajas(g(row, "bultos_servidos")),
      up(g(row, "proveedor"))
    ]);
  }
  if (!lineas.length) { console.log(`❌ ${archivo}: sin líneas`); continue; }

  const id = `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;
  if (ids.has(id)) { console.log(`⚠ ${archivo}: duplicado`); continue; }
  ids.add(id);
  despachos.push({ id, numero, archivo, tienda, fecha, hora, lanzado, finPreparacion, lineas });

  // control: unidades
  const s = lineas.reduce((a, l) => ({ p: a.p + l[4], an: a.an + l[5], sv: a.sv + l[6], bp: a.bp + l[10], bs: a.bs + l[11] }), { p: 0, an: 0, sv: 0, bp: 0, bs: 0 });
  const raro = s.an > s.p || s.sv > s.p ? "  ⚠ REVISAR" : "";
  console.log(`✔ ${archivo} · ${tienda} · ${lanzado} · ${lineas.length} lín · u ped ${s.p} anul ${s.an} serv ${s.sv} · cjs ${s.bs}/${s.bp}${raro}`);
  if (raro) console.log("   encabezados:", JSON.stringify(H.map((h, i) => h && `${XLSX.utils.encode_col(i)}:${h}`).filter(Boolean)));
}

despachos.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
fs.writeFileSync("data.json", JSON.stringify({ generado: new Date().toISOString(), columnas, despachos }));
console.log(`\nJSON generado con ${despachos.length} despachos`);
