const Excel = require("exceljs");
const fs = require("fs");

const workbook = new Excel.Workbook();
workbook.xlsx
  .readFile("humanos.xlsx")
  .then(() => {
    const worksheet = workbook.getWorksheet(1);
    const headers = worksheet.getRow(1).values.slice(1);

    const jsonArray = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      // Itera sobre cada fila
      if (rowNumber === 1) return;

      const rowValues = row.values
        .slice(1)
        .map((value) => (value === "null" || value === "" ? null : value));
      const jsonObject = {};

      headers.forEach((header, index) => {
        // Itera sobre cada columna
        jsonObject[header] = rowValues[index];
      });
      /* -------------------------------- domicilio ------------------------------- */
      const domicilio = `${jsonObject["DOMICILIO1"]} ${jsonObject["DOMICILIO2"]}`; // DOMICILIO1 y DOMICILIO2 MERGE

      jsonObject["DOMICLIO"] = domicilio;
      if (
        domicilio.includes("null null") ||
        domicilio.includes("null undefined")
      ) {
        jsonObject["DOMICLIO"] = null;
      }
      delete jsonObject["DOMICILIO1"];
      delete jsonObject["DOMICILIO2"];

      /* -------------------------------- profesion ------------------------------- */
      const profesion = ` ${jsonObject["PROFESION"]} ${jsonObject["PROFESION2"]}`; // PROFESION y PROFESION2 MERGE
      jsonObject["PROFES"] = profesion;
      if (
        profesion.includes("null null") ||
        profesion.includes("null undefined")
      ) {
        jsonObject["PROFES"] = null;
      }
      delete jsonObject["PROFESION"];
      delete jsonObject["PROFESION2"];

      /* -------------------------------- nombre ------------------------------- */
      if (jsonObject["NOMBRE"]) {
        const nombreParts = jsonObject["NOMBRE"].split(" "); // Divide el nombre en partes
        jsonObject["APE_PAT"] = nombreParts[0] || null;
        jsonObject["APE_MAT"] = nombreParts[1] || null;
        jsonObject["NOMBRES"] = nombreParts.slice(2).join(" ") || null;
        delete jsonObject["NOMBRE"];
      } else {
        jsonObject["APE_PAT"] = null;
        jsonObject["APE_MAT"] = null;
        jsonObject["NOMBRES"] = null;
        delete jsonObject["NOMBRE"];
      }

      jsonArray.push(jsonObject); // Agrega el objeto al array jsonArray
    });

    fs.writeFileSync("humanos.json", JSON.stringify(jsonArray, null, 2)); // Escribe el archivo humanos.json
  })
  .catch((err) => {
    console.error("Error leyendo el archivo:", err);
  });
