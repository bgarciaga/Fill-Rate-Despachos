const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const carpeta = "./excels";

const despachos = [];

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

const archivos = fs
  .readdirSync(carpeta)
  .filter(f => f.endsWith(".xlsx"));

for (const archivo of archivos) {

  const workbook = XLSX.readFile(
    path.join(carpeta, archivo)
  );

  const hoja =
    workbook.Sheets["Fillrate por despacho"];

  const data = XLSX.utils.sheet_to_json(
    hoja,
    {
      header: 1,
      defval: ""
    }
  );

  const tienda =
    String(data[1]?.[0] || "").trim();

  const fechaLanzado =
    String(data[2]?.[6] || "");

  const lineas = [];

  for (let i = 4; i < data.length; i++) {

    const row = data[i];

    if (!row[1]) continue;

    lineas.push([
      row[1],   // pedido
      row[3],   // codigo
      row[7],   // ean
      row[8],   // descripcion
      row[12],  // u_pedidas
      row[13],  // u_anuladas
      row[14],  // u_servidas
      row[15],  // pte_servir
      row[17],  // stock
      row[21],  // ubicacion
      row[10],  // formatos pedidos
      row[11]   // formatos servidos
    ]);
  }

  despachos.push({
    archivo,
    tienda,
    fechaLanzado,
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
  "data.json generado correctamente"
);
