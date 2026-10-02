const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const carpeta = "./excels";

const archivos = fs
  .readdirSync(carpeta)
  .filter(a => a.endsWith(".xlsx"));

const resultado = [];

for (const archivo of archivos) {

  const ruta = path.join(carpeta, archivo);

  const workbook = XLSX.readFile(ruta);

  const hoja =
    workbook.Sheets["Fillrate por despacho"];

  const data = XLSX.utils.sheet_to_json(
    hoja,
    {
      header: 1,
      defval: ""
    }
  );

  resultado.push({
    archivo,
    tienda: data[1]?.[0],
    fechaLanzado: data[1]?.[5],
    fechaFin: data[1]?.[9]
  });
}

fs.writeFileSync(
  "data.json",
  JSON.stringify(resultado, null, 2)
);

console.log("JSON generado correctamente");
