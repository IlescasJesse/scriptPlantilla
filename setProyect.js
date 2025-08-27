const mysql = require("mysql2/promise");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");

const databaseConfig = {
  host: "localhost",
  user: "root",
  password: "",
  database: "sirh",
};

async function setProyect() {
  try {
    console.log("Iniciando conexión a la base de datos...");
    const connection = await mysql.createConnection(databaseConfig);
    console.log("Conexión a la base de datos establecida.");

    const filePath = path.join(__dirname, "proyectos.xlsx");
    console.log(`Verificando existencia del archivo: ${filePath}`);
    if (!fs.existsSync(filePath)) {
      throw new Error("El archivo proyectos.xlsx no existe.");
    }
    console.log("Archivo proyectos.xlsx encontrado.");

    console.log("Leyendo archivo Excel...");
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    console.log("Convirtiendo datos de Excel a JSON...");
    const jsonData = xlsx.utils.sheet_to_json(sheet);
    console.log("Datos convertidos a JSON:", jsonData);

    for (let level = 1; level <= 5; level++) {
      const tableName = `adsc_level${level}`;
      console.log(`Procesando tabla: ${tableName}`);

      const updatePromises = jsonData.map(async (row) => {
        const { ADSCRIPCION, CLAVE, PROYECTO } = row;

        if (!ADSCRIPCION || !CLAVE || !PROYECTO) {
          console.log(
            `Datos incompletos en la fila: ADSCRIPCION=${ADSCRIPCION}, CLAVE=${CLAVE}, PROYECTO=${PROYECTO}. Saltando esta fila.`
          );
          return;
        }

        console.log(
          `Verificando existencia de ADSCRIPCION: ${ADSCRIPCION} en ${tableName}`
        );
        const [results] = await connection.execute(
          `SELECT nombre FROM ${tableName} WHERE nombre = ?`,
          [ADSCRIPCION]
        );

        if (results.length > 0) {
          console.log(
            `Actualizando registro en ${tableName} para ADSCRIPCION: ${ADSCRIPCION}`
          );
          const updateQuery = `UPDATE ${tableName} SET clave = ?, proyecto = ? WHERE nombre = ?`;
          return connection.execute(updateQuery, [
            CLAVE,
            PROYECTO,
            ADSCRIPCION,
          ]);
        } else {
          console.log(
            `No se encontró ADSCRIPCION: ${ADSCRIPCION} en ${tableName}`
          );
        }
      });

      await Promise.all(updatePromises);
      console.log(`Procesamiento de tabla ${tableName} completado.`);
    }

    console.log(
      "Conexión a MySQL establecida correctamente para modificar proyectos."
    );
  } catch (error) {
    console.error("Error exportando datos:", error);
  }
}

module.exports = { setProyect };
