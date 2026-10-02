const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const carpeta = "./excels";

const columnas = [
  "pedido",
  "codigo",
  "ean",
  "descripcion",
  "u_pedidas",
  "u_anuladas",
  "u_servidas",
  "pte_servir",
  "stock",
  "ubicacion",
  "bultos_pedidos",
  "bultos_servidos"
];

const despachos = [];

const archivos = fs
  .readdirSync(carpeta)
  .filter(f => f.toLowerCase().endsWith(".xlsx"));

for (const archivo of archivos) {

  const ruta = path.join(carpeta, archivo);

  const workbook = XLSX.readFile(ruta);

  const hoja =
    workbook.Sheets["Fillrate por despacho"];

  if (!hoja) {
    console.log(`No existe la hoja en ${archivo}`);
    continue;
  }

  const data = XLSX.utils.sheet_to_json(
    hoja,
    {
      header: 1,
      defval: ""
    }
  );

  const tienda =
    String(data?.[1]?.[0] || "")
      .replace(/\s+/g, " ")
      .trim();

  const numero =
    archivo.match(/(\d+)\.xlsx$/)?.[1] || "";

  let fecha = "";
  let hora = "";

  try {

    const fechaLanzado =
      String(data?.[2]?.[6] || "").trim();

    if (fechaLanzado) {

      const d = new Date(fechaLanzado);

      if (!isNaN(d)) {

        fecha =
          d.toISOString().substring(0, 10);

        hora =
          d.toISOString().substring(11, 16);
      }
    }

  } catch (e) {}

  const id =
    `${tienda}_${fecha}_${hora.replace(":", "")}_${numero}`;

  const lineas = [];

  for (let i = 4; i < data.length; i++) {

    const row = data[i];

    if (!row || !row[1]) continue;

    lineas.push([
      row[1] || "",     // pedido
      row[3] || "",     // codigo
      row[7] || "",     // ean
      row[8] || "",     // descripcion
      row[12] || 0,     // u_pedidas
      row[13] || 0,     // u_anuladas
      row[14] || 0,     // u_servidas
      row[15] || 0,     // pte_servir
      row[17] || 0,     // stock
      row[24] || "",    // ubicacion
      row[10] || 0,     // bultos pedidos
      row[11] || 0      // bultos servidos
    ]);
  }

  despachos.push({
    id,
    numero,
    archivo,
    tienda,
    fecha,
    hora,
    lineas
  });

}

const jsonFinal = {
  generado: new Date().toISOString(),
  columnas,
  despachos
};

fs.writeFileSync(
  "data.json",
  JSON.stringify(
    jsonFinal,
    null,
    2
  )
);

console.log(
  `JSON generado con ${despachos.length} despachos`
);
