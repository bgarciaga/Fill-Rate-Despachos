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

const COL_DEFAULT = {
  pedido: 1, codigo: 3, ean: 7, descripcion: 8,
  bultos_pedidos: 10, bultos_servidos: 11,
  u_pedidas: 12, u_anuladas: 13, u_servidas: 14, pte_servir: 15,
  stock: 17, ubicacion: 24, proveedor: -1
};

const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
const num = v => { if (typeof v === "number") return v; const n = parseFloat(String(v ?? "").replace(/,/g, "").trim()); return isNaN(n) ? 0 : n; };
const isNum = v => typeof v === "number" || /^-?[\d,]+(\.\d+)?$/.test(String(v ?? "").trim());
const txt = v => String(v ?? "").replace(/\s+/g, " ").trim();
const pad = n => String(n).padStart(2, "0");
const cleanId = v => txt(v).replace(/\.0+$/, "");

// ---------- FECHAS ----------
function parseFecha(v) {
  if (v === "" || v == null) return null;
  if (typeof v === "number") {
    if (v < 44000 || v > 50000) return null;
    const p = XLSX.SSF.parse_date_code(v);
    return p ? { fecha: `${p.y}-${pad(p.m)}-${pad(p.d)}`, hora: `${pad(p.H)}:${pad(p.M)}`, seg: pad(Math.floor(p.S)) } : null;
  }
  const s = txt(v);
  let m = s.match(/(20\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) { const [, y, mo, d, H = "0", M = "0", S = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}`, seg: pad(S) }; }
  m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](20\d{2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) { const [, d, mo, y, H = "0", M = "0", S = "0"] = m; return { fecha: `${y}-${pad(mo)}-${pad(d)}`, hora: `${pad(H)}:${pad(M)}`, seg: pad(S) }; }
  return null;
}

// Todas las fechas en las primeras filas (antes de los encabezados), en orden
function fechasEncabezado(data, hasta) {
  const out = [];
  for (let r = 0; r < Math.min(hasta, data.length); r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      const f = parseFecha(row[c]);
      if (f) out.push({ ...f, r, c, label: norm(row[c - 1] || (data[r - 1] || [])[c] || "") });
    }
  }
  return out;
}

// ---------- COLUMNAS ----------
const ffill = row => { let last = ""; return (row || []).map(v => (txt(v) ? (last = txt(v)) : last)); };

function detectarColumnas(data) {
  let hr = -1;
  for (let r = 0; r < Math.min(15, data.length); r++) {
    if ((data[r] || []).some(c => norm(c).includes("descripcion"))) { hr = r; break; }
  }
  const idx = { ...COL_DEFAULT };
  if (hr < 0) return { hr: 4, inicio: 4, idx, labels: [] };

  const sub = data[hr + 1] || [];
  const subEsHeader = sub.some(c => /[a-z]/i.test(String(c))) && !sub.some(c => typeof c === "number");
  const ancho = Math.max(...data.slice(Math.max(0, hr - 1), hr + 3).map(r => (r || []).length));
  const top = ffill(data[hr - 1]), main = data[hr] || [], mainF = ffill(main);

  const labels = [];
  for (let c = 0; c < ancho; c++) {
    labels[c] = subEsHeader
      ? norm(`${mainF[c]} ${txt(sub[c])}`)
      : norm(`${top[c] && norm(top[c]) !== norm(main[c]) ? top[c] + " " : ""}${txt(main[c])}`);
  }
  const inicio = hr + (subEsHeader ? 2 : 1);

  const rows = data.slice(inicio).filter(r => r && r.filter(v => v !== "").length > 4).slice(0, 300);
  const stat = c => {
    const vals = rows.map(r => r[c]).filter(v => v !== "" && v != null);
    const n = vals.length || 1;
    return {
      n: vals.length,
      ean: vals.filter(v => /^\d{8,14}$/.test(cleanId(v))).length / n,
      ubic: vals.filter(v => /^[A-Z0-9]{1,4}(-[A-Z0-9]{1,4}){2,}$/i.test(txt(v))).length / n,
      len: vals.reduce((a, v) => a + (/[a-z]/i.test(String(v)) ? txt(v).length : 0), 0) / n
    };
  };
  const S = labels.map((_, c) => stat(c));
  const used = new Set();
  const take = (k, c) => { if (c >= 0) { idx[k] = c; used.add(c); } };
  const byLabel = test => labels.findIndex((l, c) => !used.has(c) && l && test(l));
  const best = (score, min) => {
    let bi = -1, bv = min;
    S.forEach((s, c) => { if (!used.has(c) && s.n && score(s) > bv) { bv = score(s); bi = c; } });
    return bi;
  };
  const isB = l => /bult|bto|caja/.test(l);

  take("proveedor", byLabel(l => /proveedor/.test(l)));
  take("pedido", byLabel(l => /pedido/.test(l) && !isB(l) && !/pedid[ao]s/.test(l)));
  take("codigo", byLabel(l => /codigo|articulo|sku/.test(l) && !/barra|ean/.test(l)));
  const eanL = byLabel(l => /\bean\b|barra/.test(l));
  take("ean", eanL >= 0 ? eanL : best(s => s.ean, 0.6));
  take("descripcion", byLabel(l => /descripcion/.test(l)));
  if (S[idx.descripcion] && S[idx.descripcion].len < 5) { used.delete(idx.descripcion); take("descripcion", best(s => s.len, 10)); }
  take("ubicacion", best(s => s.ubic, 0.5));
  take("stock", byLabel(l => /stock|existencia/.test(l)));
  take("bultos_pedidos", byLabel(l => isB(l) && /ped/.test(l)));
  take("bultos_servidos", byLabel(l => isB(l) && /serv/.test(l)));
  take("u_anuladas", byLabel(l => !isB(l) && /anul/.test(l)));
  take("pte_servir", byLabel(l => !isB(l) && /pte|pend/.test(l)));
  take("u_servidas", byLabel(l => !isB(l) && /serv/.test(l)));
  take("u_pedidas", byLabel(l => !isB(l) && /pedid|ped\b|u\.? ?ped|cant/.test(l)));

  return { hr, inicio, idx, labels };
}

// ---------- PROCESO ----------
if (!fs.existsSync(carpeta)) { console.log("❌ La carpeta no existe:", path.resolve(carpeta)); process.exit(1); }

const archivos = fs.readdirSync(carpeta).filter(f => /\.xlsx?$/i.test(f) && !f.startsWith("~$")).sort();
console.log("Archivos Excel:", archivos.length);

const despachos = [];
const ids = new Set();

for (const archivo of archivos) {
  const ruta = path.join(carpeta, archivo);
  console.log(`\n=== ${archivo} ===`);

  let wb;
  try { wb = XLSX.readFile(ruta); } catch (e) { console.log("❌ No se pudo leer:", e.message); continue; }
  const nombreHoja = wb.SheetNames.find(n => norm(n) === norm(HOJA))
    || wb.SheetNames.find(n => /fill ?rate/.test(norm(n))) || wb.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { header: 1, defval: "", raw: true });
  data.slice(0, 6).forEach((r, i) => console.log(`  fila ${i}:`, JSON.stringify(r)));

  const { hr, inicio, idx, labels } = detectarColumnas(data);

  const tienda = (txt(data?.[1]?.[0]) || archivo.replace(/-\d+\.xlsx?$/i, "")).toUpperCase();
  const numero = archivo.match(/(\d+)\.xlsx?$/i)?.[1] || "";

  // fechas del encabezado: la que tenga etiqueta "lanz" o la primera = lanzado; "fin" o la segunda = fin preparación
  const fechas = fechasEncabezado(data, Math.max(hr, 4));
  console.log("Fechas encontradas:", JSON.stringify(fechas.map(f => `${f.fecha} ${f.hora} (fila ${f.r}, col ${f.c}, "${f.label}")`)));
  let fl = fechas.find(f => /lanz/.test(f.label)) || fechas[0];
  let fp = fechas.find(f => /fin|final|termin/.test(f.label)) || fechas.find(f => f !== fl);
  if (fp && fl && (fp.fecha + fp.hora) < (fl.fecha + fl.hora)) [fl, fp] = [fp, fl];

  if (!fl) {
    const m = fs.statSync(ruta).mtime;
    fl = { fecha: `${m.getFullYear()}-${pad(m.getMonth() + 1)}-${pad(m.getDate())}`, hora: `${pad(m.getHours())}:${pad(m.getMinutes())}`, seg: "00" };
    console.log("⚠ Sin fecha de lanzado, uso fecha del archivo");
  }
  const { fecha, hora } = fl;
  const lanzado = `${fecha.replace(/-/g, "/")} ${hora}:${fl.seg || "00"}`;
  const finPreparacion = fp ? `${fp.fecha} ${fp.hora}` : "";

  console.log("Encabezados:", JSON.stringify(labels.map((l, i) => l && `${i}:${l}`).filter(Boolean)));
  console.log("Columnas:", JSON.stringify(idx));

  const lineas = [];
  for (let i = inicio; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const pedido = cleanId(row[idx.pedido]);
    if (!pedido || !/^\d+$/.test(pedido)) continue;
    lineas.push([
      pedido,
      cleanId(row[idx.codigo]),
      cleanId(row[idx.ean]),
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

  // Bultos servidos vacíos → calcular proporcional a unidades servidas
  const sumBS = lineas.reduce((a, l) => a + l[11], 0);
  const sumUS = lineas.reduce((a, l) => a + l[6], 0);
  if (!sumBS && sumUS) {
    console.log("⚠ Bultos servidos vacíos, se calculan por unidades");
    lineas.forEach(l => { l[11] = l[4] ? Math.round(l[10] * l[6] / l[4]) : 0; });
  }

  const ok = lineas.filter(l => Math.abs(l[7] - (l[4] - l[6])) < 0.01).length;
  console.log(`Tienda: ${tienda} · lanzado ${lanzado} · fin ${finPreparacion || "—"} · líneas ${lineas.length} · cuadran ${ok}/${lineas.length}`);
  lineas.slice(0, 2).forEach(l => console.log("  ", JSON.stringify(l)));
  if (!lineas.length) { console.log("❌ Sin líneas, se omite"); continue; }

  const id = `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;
  if (ids.has(id)) { console.log("❌ Duplicado, se omite"); continue; }
  ids.add(id);

  despachos.push({ id, numero, archivo, tienda, fecha, hora, lanzado, finPreparacion, lineas });
}

despachos.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
fs.writeFileSync("data.json", JSON.stringify({ generado: new Date().toISOString(), columnas, despachos }));
console.log(`\nJSON generado con ${despachos.length} despachos`);
