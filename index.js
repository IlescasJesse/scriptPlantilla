const { MongoClient } = require("mongodb");
const Excel = require("exceljs");
const fs = require("fs");
const { ObjectId } = require("mongodb");
const { actualizarPlantillaDesdeMongo } = require("./comisionados");
const {
  procesarPlantillatipoNOM,
  actualizarTiponomEnPlazas,
} = require("./tiponomina");
const { setProyect } = require("./setProyect");
const { exec } = require("child_process");
const { json } = require("stream/consumers");
// const uri = "mongodb://mongoadmin:pb9*V82nY3An@172.17.90.58:3001";
const uri = "mongodb://admin:1234@localhost:27017/";
const client = new MongoClient(uri);

async function run() {
  try {
    console.log("Connecting to MongoDB...");
    await client.connect();
    const database = client.db("sirhTest");
    const collectionsToDelete = [
      "BAJAS",
      "BITACORA",
      "INCAPACIDADES",
      "JUSTIFICACIONES",
      "PERMISOS_ECONOMICOS",
      "PERMISOS_EXT",
      "PLANTILLA",
      "PLAZAS",
      "LICENCIAS",
      "INCIDENCIAS",
      "INASISTENCIAS",
      "HSY_LICENCIAS",
      "HSY_RECATEGORIZACIONES",
      "HSY_PROYECTOS",
      "HSY_STATUS_EMPLEADO",
      "USERS_ACTIONS",
      "USER_ACTIONS",
      "PER_VACACIONALES_BASE",
      "PER_VACACIONALES_CONTRATO",
      "PLANTILLA_FORANEA",
      "TALONES",
    ];

    console.log("Deleting specified collections...");
    for (const collectionName of collectionsToDelete) {
      const collection = database.collection(collectionName);
      const exists = await database
        .listCollections({ name: collectionName })
        .hasNext();
      if (exists) {
        await collection.drop();
        console.log(`Collection ${collectionName} deleted successfully.`);
      } else {
        console.log(`Collection ${collectionName} does not exist.`);
      }
    }

    console.log("All specified collections processed for deletion.");

    const collectionPlantilla = database.collection("PLANTILLA");

    const collectionBitacora = database.collection("BITACORA");

    // LEEMOS LA PLANTILLA
    console.log("Reading plantilla_test.xlsx...");
    const workbookPlantilla = new Excel.Workbook();
    await workbookPlantilla.xlsx.readFile("plantilla_2026_test.xlsx");
    const worksheetPlantilla = workbookPlantilla.getWorksheet(1);
    const headersPlantilla = worksheetPlantilla.getRow(1).values.slice(1);

    // LEEMOS COMISIONADOS
    console.log("Reading COMISIONADOS_SINDICALES.xlsx...");
    const workbookVacaciones = new Excel.Workbook();
    await workbookVacaciones.xlsx.readFile("VACACIONES.xlsx");
    const worksheetVacaciones = workbookVacaciones.getWorksheet("vacaciones");

    const workbookComisionados = new Excel.Workbook();
    await workbookComisionados.xlsx.readFile("COMISIONADOS_SINDICALES.xlsx");
    const worksheetComisionados =
      workbookComisionados.getWorksheet("comisionados");

    console.log("Extracting names from COMISIONADOS_SINDICALES.xlsx...");

    console.log("Names extracted from COMISIONADOS_SINDICALES.xlsx.");

    // LEEMOS TARJETAS
    console.log("Reading TARJETAS.xlsx...");
    const workbookTarjetas = new Excel.Workbook();
    await workbookTarjetas.xlsx.readFile("TARJETAS.xlsx");

    const tarjetasData = [];
    workbookTarjetas.eachSheet((worksheet) => {
      const nombreIndex = 3; // Columna "NOMBRE"
      const numIndex = 1; // Columna "NUM"

      // Procesar las filas
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1) return; // Saltar encabezados
        const nombre = row.getCell(nombreIndex).value;
        const numTarjeta = row.getCell(numIndex).value;
        if (nombre && numTarjeta) {
          tarjetasData.push({
            nombre: nombre
              .trim()
              .replace(/\s{2,}/g, " ")
              .replace(/\.$/, "")
              .toUpperCase(),
            numTarjeta: numTarjeta,
            horario: row.getCell(5).value,
          });
        }
      });

      // Procesar las filas
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1) return; // Saltar encabezados
        const nombre = row.getCell(nombreIndex).value;
        const numTarjeta = row.getCell(numIndex).value;
        if (nombre && numTarjeta) {
          tarjetasData.push({
            nombre: nombre.trim().toUpperCase(),
            numTarjeta: numTarjeta,
            horario: row.getCell(5).value,
          });
        }
      });
    });

    console.log("Processing rows from plantilla_2026_test.xlsx...");
    const jsonArray = [];
    const licenciaArray = [];

    worksheetPlantilla.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const rowValues = row.values
        .slice(1)
        .map((value) => (value === "null" || value === "" ? null : value));
      const jsonObject = {};

      headersPlantilla.forEach((header, index) => {
        jsonObject[header] = rowValues[index];
      });
      jsonObject["ID_CTRL_ASIST"] = new ObjectId();
      while (
        jsonArray.some((item) =>
          item["ID_CTRL_ASIST"].equals(jsonObject["ID_CTRL_ASIST"])
        )
      ) {
        jsonObject["ID_CTRL_ASIST"] = new ObjectId();
      }
      // Asignar correctamente el domicilio combinando DOMICILIO1 y DOMICILIO2 si existen
      const domicilio1 = jsonObject["DOMICILIO1"]
        ? jsonObject["DOMICILIO1"].toString().trim()
        : "";
      const domicilio2 = jsonObject["DOMICILIO2"]
        ? jsonObject["DOMICILIO2"].toString().trim()
        : "";
      let domicilio = domicilio1;
      if (domicilio2 && domicilio2 !== domicilio1) {
        domicilio =
          domicilio2 && domicilio1
            ? `${domicilio1} ${domicilio2}`
            : domicilio1 || domicilio2 || " ";
      }
      jsonObject["DOMICILIO"] = domicilio.length > 0 ? domicilio : null;
      delete jsonObject["DOMICILIO1"];
      delete jsonObject["DOMICILIO2"];

      jsonObject["ADSCRIPCION"] = jsonObject["DEPARTAMENTO"];

      // Convert FECHA_INGRESO from DD/MM/YYYY to YYYY/MM/DD
      if (
        jsonObject["FECHA_INGRESO"] &&
        typeof jsonObject["FECHA_INGRESO"] === "string" &&
        jsonObject["FECHA_INGRESO"].includes("/")
      ) {
        const [day, month, year] = jsonObject["FECHA_INGRESO"].split("/");
        jsonObject["FECHA_INGRESO"] = `${year}/${month}/${day}`;
      }
      const profesion2 =
        jsonObject["PROFESION2"] === null ? " " : jsonObject["PROFESION2"];
      const profesion1 =
        jsonObject["PROFESION"] === null ? " " : jsonObject["PROFESION"];
      const profesion = `${profesion1} ${profesion2}`;

      jsonObject["PROFES"] = profesion === "   " ? null : profesion;

      // Remove everything after "." in TURNOMAT
      if (
        jsonObject["TURNOMAT"] &&
        typeof jsonObject["TURNOMAT"] === "string"
      ) {
        jsonObject["TURNOMAT"] = jsonObject["TURNOMAT"].split(".")[0];
      }
      if (
        jsonObject["CURP"] &&
        typeof jsonObject["CURP"] === "string" &&
        jsonObject["CURP"].length >= 10
      ) {
        const curp = jsonObject["CURP"].toUpperCase();
        // CURP positions 5-10: YYMMDD
        const fechaCurp = curp.substring(4, 10);
        let year = fechaCurp.substring(0, 2);
        const month = fechaCurp.substring(2, 4);
        const day = fechaCurp.substring(4, 6);

        // If year >= 00 and <= current year (e.g. 24), assume 2000+, else 1900+
        const currentYear = new Date().getFullYear() % 100;
        year = parseInt(year, 10) <= currentYear ? `20${year}` : `19${year}`;

        jsonObject["FECHA_NAC"] = `${year}/${month}/${day}`;
      } else {
        jsonObject["FECHA_NAC"] = null;
      }
      delete jsonObject["PROFESION"];
      delete jsonObject["PROFESION2"];

      if (
        jsonObject["NOMBRE"] !== "V A C A N T E DE:" &&
        jsonObject["NOMBRE"] !== null
      ) {
        const nombreParts = jsonObject["NOMBRE"].split(" ");
        jsonObject["APE_PAT"] = nombreParts[0] || null;
        jsonObject["APE_MAT"] = nombreParts[1] || null;
        jsonObject["NOMBRES"] = nombreParts.slice(2).join(" ") || null;

        // Compare with TARJETAS.xlsx
        const tarjetaMatch = tarjetasData.find(
          (tarjeta) => tarjeta.nombre === jsonObject["NOMBRE"]
        );
        jsonObject["NUMTARJETA"] = tarjetaMatch
          ? tarjetaMatch.numTarjeta
          : null;

        delete jsonObject["NOMBRE"];
      } else {
        jsonObject["APE_PAT"] = null;
        jsonObject["APE_MAT"] = null;
        jsonObject["NOMBRES"] = null;
        jsonObject["STATUS_EMPLEADO"] = null;
        jsonObject["NUMTARJETA"] = null;
        delete jsonObject["NOMBRE"];
      }

      // Process other fields (e.g., licencia, etc.)
      let licencia1 =
        jsonObject["LICENCIA"] === null || jsonObject["LICENCIA"] === undefined
          ? "  "
          : jsonObject["LICENCIA"];
      const licencia2 =
        jsonObject["LICENCIA1"] === null ||
        jsonObject["LICENCIA1"] === undefined
          ? " "
          : jsonObject["LICENCIA1"];
      delete jsonObject["LICENCIA"];
      delete jsonObject["LICENCIA1"];
      delete jsonObject["SUELDO_GRV"];

      delete jsonObject["GUARDE"];
      delete jsonObject["GASCOM"];
      const numpla = jsonObject["NUMPLA"] === null ? " " : jsonObject["NUMPLA"];

      const TIPONOM =
        jsonObject["TIPONOM"] === null ? " " : jsonObject["TIPONOM"];

      if (licencia1 && licencia1.startsWith("CUBRE A:")) {
      } else if (licencia1 && licencia1.startsWith("SUST. A:")) {
        licencia1 = licencia1.split(":")[1]
          ? licencia1.split(":")[1].trim()
          : "";
      } else if (licencia1 && licencia1.startsWith("SUST. A")) {
        licencia1 = licencia1.split("A")[1]
          ? licencia1.split("A")[1].trim()
          : "";
        licencia1 = licencia1.split(":")[1];
      }

      const licenciaObject = {
        previousOcuppants: [
          {
            NOMBRE: licencia1,
            FECHA: null,
            FECHA_BAJA: null,
            MOTIVO_BAJA: licencia2,
          },
        ],
        NUMPLA: numpla,
        TIPONOM: TIPONOM,
      };
      const proyecto =
        jsonObject["PROYECTO"] === null ? " " : jsonObject["PROYECTO"];
      const departamento =
        jsonObject["DEPARTAMENTO"] === null ? " " : jsonObject["DEPARTAMENTO"];

      licenciaObject["PROYECTO"] = proyecto;

      licenciaObject["DEPARTAMENTO"] = departamento;
      jsonObject["ID_CTRL_ASSIST"] = new ObjectId();
      jsonObject["ID_CTRL_TALON"] = new ObjectId();
      jsonObject["ID_CTRL_NOM"] = new ObjectId();
      jsonObject["ID_CTRL_CAP"] = new ObjectId();
      // Buscar coincidencia en worksheetVacaciones por NUMPLA
      let vacacionesMatch = null;
      worksheetVacaciones.eachRow(
        { includeEmpty: true },
        (vacRow, vacRowNumber) => {
          if (vacRowNumber === 1) return; // Saltar encabezados
          const NUE = parseInt(vacRow.getCell(3).value, 10); // Ajusta el índice si es diferente
          if (NUE === jsonObject["NUMEMP"]) {
            let fechaVac = vacRow.getCell(4).value || null;
            if (
              fechaVac &&
              typeof fechaVac === "string" &&
              fechaVac.includes("/")
            ) {
              const [day, month, year] = fechaVac.split("/");
              fechaVac = `${year}/${month}/${day}`;
            }
            vacacionesMatch = {
              PERIODO: 0, // Ajusta el índice si es diferente
              FECHA_VACACIONES: fechaVac,
            };
          }
        }
      );
      jsonObject["VACACIONES"] = vacacionesMatch || {
        PERIODO: 0,
        FECHA_VACACIONES: null,
      };

      if (
        (jsonObject["NOMBRES"] && jsonObject["NOMBRES"].includes("VACANTE")) ||
        jsonObject["APE_PAT"] === null ||
        jsonObject["APE_MAT"] === null
      ) {
        licenciaObject["status"] = 2;
        jsonObject["status"] = 2;
      } else {
        licenciaObject["status"] = 1;
        jsonObject["status"] = 1;
      }
      if (
        [
          "1140041480100000220",
          "1140041480100000222",
          "1140041480100000223",
          "1140041480100000227",
          "1140041480100000226",
        ].includes(jsonObject["PROYECTO"])
      ) {
        jsonObject["AREA_RESP"] = "AUD";
      } else if (
        ["1140051490100000500", "1140051490100000508"].includes(
          jsonObject["PROYECTO"]
        )
      ) {
        jsonObject["AREA_RESP"] = "PLAN";
      } else {
        jsonObject["AREA_RESP"] = "CTRAL";
      }

      licenciaArray.push(licenciaObject);
      jsonArray.push(jsonObject);
    });

    console.log("Inserting documents into MongoDB...");

    // Insert documents into PLANTILLA collection
    const resultPlantilla = await collectionPlantilla.insertMany(jsonArray);

    // Crear objetos bitácora después de insertar en PLANTILLA_2025
    const bitacoraArray = Object.values(resultPlantilla.insertedIds).map(
      (id, index) => {
        const empleado = jsonArray[index];
        const personalEntries = [
          {
            autor: "SISTEMA",
            comentario: "GENERACIÓN DE PLANTILLA",
            fecha: new Date(),
          },
        ];

        // Si la CLAVE es 105, agregar comentario con datos de MADRE y PADRE
        if (empleado["CLAVE"] === 105) {
          const madre = empleado["MADRE"] || "No especificada";
          const padre = empleado["PADRE"] || "No especificado";
          personalEntries.push({
            autor: "SISTEMA",
            comentario: `MADRE: ${madre} | PADRE: ${padre}`,
            fecha: new Date(),
          });
        }

        return {
          personal: personalEntries,
          incidencias: [],
          nomina: [],
          archivo: [],
          tramites: [],
          capacitaciones: [],
          id_plantilla: id,
          vacaciones: [],
          talon: [],
        };
      }
    );
    const resultBitacora = await collectionBitacora.insertMany(bitacoraArray);
    // --- Inicio: crear colección TALONES (un documento por empleado con array TALONES de 2 items para diciembre) ---
    console.log("Creating TALONES collection...");
    const collectionTalones = database.collection("TALONES");

    // Fechas de pago solo para diciembre 2025 (quincenas 23 y 24)
    const fechasPagoDiciembre = {
      23: "2025-12-15", // Primera quincena de diciembre (1-15)
      24: "2025-12-31", // Segunda quincena de diciembre (16-31)
    };

    const talonesArray = Object.values(resultPlantilla.insertedIds).map(
      (id) => {
        // Crear array de solo 2 talones (quincenas 23 y 24 de diciembre)
        const talones = [];
        for (let quin = 23; quin <= 24; quin++) {
          talones.push({
            _id: new ObjectId(),
            QUIN: quin,
            FECHA_PAG: fechasPagoDiciembre[quin],
            STATUS: 2,
            FOLIO: null,
          });
        }

        return {
          _id: new ObjectId(),
          _idEmployee: id,
          TALONES: talones,
        };
      }
    );

    await collectionTalones.insertMany(talonesArray);
    console.log(
      `${talonesArray.length} documents inserted into TALONES collection (only December quincenas 23 & 24).`
    );

    // Crear índices útiles
    await collectionTalones.createIndex({ _idEmployee: 1 });
    await collectionTalones.createIndex({ "TALONES.QUIN": 1 });

    // --- Fin: TALONES ---
    const permisos_economicos = [];
    const incapacidades = [];
    const vacaciones = [];
    const eximas = [];
    const collectionPermisosEconomicos = database.collection(
      "PERMISOS_ECONOMICOS"
    );
    const collectionIncapacidades = database.collection("INCAPACIDADES");
    const collectionEximas = database.collection("EXIMAS");
    // Crear colección VACACIONES_BASE con 6 documentos
    const collectionVacacionesBase = database.collection(
      "PER_VACACIONALES_BASE"
    );
    const collectionVacacionesContrato = database.collection(
      "PER_VACACIONALES_CONTRATO"
    );

    // Crear colección PER_VACACIONALES_BASE con 6 documentos
    // Duplicate declaration removed. The previous vacacionesBaseDocs and insertMany already exist above.

    // Crear colección PER_VACACIONALES_CONTRATO con 8 documentos
    const vacacionesContratoDocs = [];
    for (let periodo = 0; periodo <= 7; periodo++) {
      vacacionesContratoDocs.push({
        PERIODO: periodo + 1,
        10: { FECHA_INI: null, FECHA_FIN: null },
        11: { FECHA_INI: null, FECHA_FIN: null },
        12: { FECHA_INI: null, FECHA_FIN: null },
        13: { FECHA_INI: null, FECHA_FIN: null },
        14: { FECHA_INI: null, FECHA_FIN: null },
        15: { FECHA_INI: null, FECHA_FIN: null },
        16: { FECHA_INI: null, FECHA_FIN: null },
      });
    }
    await collectionVacacionesContrato.insertMany(vacacionesContratoDocs);
    const vacacionesBaseDocs = [];
    for (let periodo = 0; periodo <= 5; periodo++) {
      vacacionesBaseDocs.push({
        PERIODO: periodo + 1,
        11: { FECHA_INI: null, FECHA_FIN: null },
        13: { FECHA_INI: null, FECHA_FIN: null },
        15: { FECHA_INI: null, FECHA_FIN: null },
        17: { FECHA_INI: null, FECHA_FIN: null },
        19: { FECHA_INI: null, FECHA_FIN: null },
      });
    }
    await collectionVacacionesBase.insertMany(vacacionesBaseDocs);
    if (permisos_economicos.length > 0) {
      await collectionPermisosEconomicos.insertMany(permisos_economicos);
    }
    if (incapacidades.length > 0) {
      await collectionIncapacidades.insertMany(incapacidades);
    }
    if (vacaciones.length > 0) {
      await collectionVacaciones.insertMany(vacaciones);
    }
    if (eximas.length > 0) {
      await collectionEximas.insertMany(eximas);
    }

    console.log("Writing JSON files...");
    fs.writeFileSync("plazas.json", JSON.stringify(licenciaArray, null, 2));
    fs.writeFileSync("plantilla.json", JSON.stringify(jsonArray, null, 2));
    fs.writeFileSync("bitacora.json", JSON.stringify(bitacoraArray, null, 2));
    // Insert plazas.json into the PLAZAS collection
    const collectionPlazas = database.collection("PLAZAS");
    const plazasData = JSON.parse(fs.readFileSync("plazas.json", "utf8"));
    await collectionPlazas.insertMany(plazasData);

    // Update plantilla before inserting into MongoDB
    const bulkOpsPlantillaUpdate = jsonArray.map((item) => ({
      updateOne: {
        filter: { ID_CTRL_ASIST: item.ID_CTRL_ASIST },
        update: { $set: item },
        upsert: true,
      },
    }));

    await collectionPlantilla.bulkWrite(bulkOpsPlantillaUpdate);

    // Insert updated plantilla into MongoDB
    const plantillaPath = "plantilla.json";

    console.log("JSON files written successfully");

    // Actualizar los ids en las colecciones
    const bulkOpsPlantilla = Object.values(resultPlantilla.insertedIds).map(
      (id, index) => ({
        updateOne: {
          filter: { _id: id },
          update: {
            $set: {
              ID_BITACORA: Object.values(resultBitacora.insertedIds)[index],
            },
          },
        },
      })
    );

    const bulkOpsBitacora = Object.values(resultBitacora.insertedIds).map(
      (id, index) => ({
        updateOne: {
          filter: { _id: id },
          update: {
            $set: {
              id_plantilla: Object.values(resultPlantilla.insertedIds)[index],
            },
          },
        },
      })
    );

    await collectionPlantilla.bulkWrite(bulkOpsPlantilla);
    await collectionBitacora.bulkWrite(bulkOpsBitacora);

    console.log("Generating plantilla structure...");

    console.log("Updating PLANTILLA.json...");
    console.log("Documents inserted into MongoDB successfully");
    // setProyect();
    actualizarPlantillaDesdeMongo();
    procesarPlantillatipoNOM();
    actualizarTiponomEnPlazas();
  } catch (err) {
    console.error("Error:", err);
  } finally {
    console.log("Closing MongoDB connection...");
    await client.close();
    console.log("MongoDB connection closed");
  }
}

run().catch(console.dir);
