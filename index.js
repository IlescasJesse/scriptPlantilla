const { MongoClient } = require("mongodb");
const Excel = require("exceljs");
const fs = require("fs");
const { ObjectId } = require("mongodb");
const { actualizarPlantillaDesdeMongo } = require("./COMISIONES/comisionados");
const {
  procesarPlantillatipoNOM,
  actualizarTiponomEnPlazas,
} = require("./tiponomina");
// const uri = "mongodb://mongoadmin:pb9*V82nY3An@172.17.90.58:3001";
const uri = "mongodb://admin:1234@localhost:27017/";
const client = new MongoClient(uri);

async function run() {
  try {
    console.log("Connecting to MongoDB...");
    await client.connect();
    const database = client.db("SIRH2026");
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
    console.log("Reading plantilla_2026_test.xlsx...");
    const workbookPlantilla = new Excel.Workbook();
    await workbookPlantilla.xlsx.readFile("plantilla_2026_test.xlsx");
    const worksheetPlantilla = workbookPlantilla.getWorksheet(1);
    const headersPlantilla = worksheetPlantilla.getRow(1).values.slice(1);

    // LEEMOS VACACIONES_BASE
    console.log("Reading VACACIONES_BASE.xlsx...");
    const workbookVacacionesBase = new Excel.Workbook();
    await workbookVacacionesBase.xlsx.readFile("VACACIONES/VACACIONES_BASE.xlsx");
    const worksheetVacacionesBase = workbookVacacionesBase.getWorksheet("vacaciones");

    // LEEMOS VACACIONES_CONFIANZA
    console.log("Reading VACACIONES_CONFIANZA.xlsx...");
    const workbookVacacionesConf = new Excel.Workbook();
    await workbookVacacionesConf.xlsx.readFile("VACACIONES/VACACIONES_CONFIANZA.xlsx");
    const worksheetVacacionesConf = workbookVacacionesConf.getWorksheet("vacaciones");

    // LEEMOS TARJETAS – un único archivo con RFC/NUMTARJETA (y área, etc.)
    console.log("Reading TARJETAS.xlsx...");
    const tarjetasData = [];
    const workbookTarjetas = new Excel.Workbook();
    await workbookTarjetas.xlsx.readFile("NUMEROS_TARJETAS/TARJETAS.xlsx");
    workbookTarjetas.eachSheet((worksheet) => {
      let rfcIdx = null,
        numIdx = null,
        horM = null,
        horV = null,
        areaIdx = null;

      worksheet.getRow(1).eachCell((cell, col) => {
        const h = (cell.text || cell.value || "").toString().toUpperCase();
        if (h.includes("RFC")) rfcIdx = col;
        if (h.includes("NUMTARJETA")) numIdx = col;
        if (h.includes("TURNOMAT")) horM = col;
        if (h.includes("TURNOVES")) horV = col;
        if (h.includes("AREA_RESP")) areaIdx = col;
      });
      if (!rfcIdx || !numIdx) return;

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1) return;
        const rawRfc = row.getCell(rfcIdx).value;
        const numTarjeta = row.getCell(numIdx).value;
        if (!rawRfc || !numTarjeta) return;
        tarjetasData.push({
          rfc: rawRfc.toString().trim().toUpperCase(),
          numTarjeta,
          horarioM: horM ? row.getCell(horM).value : undefined,
          horarioV: horV ? row.getCell(horV).value : undefined,
          area: areaIdx ? row.getCell(areaIdx).value : undefined,
        });
      });
    });

    console.log("Processing rows from plantilla_2026_test.xls...");
    let jsonArray = [];
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
      const profesion2 = jsonObject["PROFESION2"] ?? " ";
      const profesion1 = jsonObject["PROFESION"] ?? " ";
      const profesion = `${profesion1} ${profesion2}`;
      jsonObject["PROFES"] = profesion.trim() === "" ? null : profesion.trim();

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

        // Actualizar el campo NUMTARJETA buscando por RFC en el archivo de tarjetas
        const normalizedRfc = jsonObject["RFC"]
          ? jsonObject["RFC"].toString().trim().toUpperCase()
          : null;
        if (normalizedRfc) {
          const tarjetaMatch = tarjetasData.find((t) => t.rfc === normalizedRfc);
          if (tarjetaMatch) {
            jsonObject["NUMTARJETA"] = tarjetaMatch.numTarjeta;
            if (tarjetaMatch.horarioM !== undefined)
              jsonObject["TURNOMAT"] = tarjetaMatch.horarioM;
            if (tarjetaMatch.horarioV !== undefined)
              jsonObject["TURNOVES"] = tarjetaMatch.horarioV;
            if (tarjetaMatch.area !== undefined)
              jsonObject["AREA_RESP"] = tarjetaMatch.area;
          } else {
            jsonObject["AREA_RESP"] = "CTRAL";
          }
        }

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
      let vacacionesMatchConfianza = null;
      worksheetVacacionesConf.eachRow(
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
            vacacionesMatchConfianza = {
              PERIODO: 0, // Ajusta el índice si es diferente
              FECHA_VACACIONES: fechaVac,
            };
          }
        }
      );

      // Buscar coincidencia en worksheetVacacionesBase por NUMEMP
      let vacacionesMatchBase = null;
      worksheetVacacionesBase.eachRow(
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
            vacacionesMatchBase = {
              PERIODO: 0, // Ajusta el índice si es diferente
              FECHA_VACACIONES: fechaVac,
            };
          }
        }
      );

      jsonObject["VACACIONES"] = vacacionesMatchConfianza ||
        vacacionesMatchBase || {
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

      licenciaArray.push(licenciaObject);
      jsonArray.push(jsonObject);
    });

    console.log("Inserting documents into MongoDB...");

    // Insert documents into PLANTILLA collection
    const resultPlantilla = await collectionPlantilla.insertMany(jsonArray);
    // keep snapshot for later bitacora creation (before we remove LS entries)
    const originalJsonArray = [...jsonArray];

    // si una NUMPLA aparece con LS y con B, vamos a:
    //   1. trasladar el documento LS a LICENCIAS
    //   2. eliminar ese LS de PLANTILLA
    //   3. (B se queda en PLANTILLA, sin cambios)
    const numplasConLS = new Set();
    const numplasConB = new Set();

    jsonArray.forEach((item) => {
      if ((item.TIPONOM || "").toUpperCase() === "LS") {
        numplasConLS.add(item.NUMPLA);
      }
      if ((item.TIPONOM || "").toUpperCase() === "B") {
        numplasConB.add(item.NUMPLA);
      }
    });

    const licCollection = database.collection("LICENCIAS");
    const licToInsert = [];
    const idsToDelete = [];

    // Para cada LS (dueño) que tenga un B (cubridor) en la misma NUMPLA
    Object.entries(resultPlantilla.insertedIds).forEach(([idx, id]) => {
      const item = jsonArray[idx];
      const tip = (item.TIPONOM || "").toUpperCase();

      if (tip === "LS" && numplasConB.has(item.NUMPLA)) {
        // Crear documento para LICENCIAS con datos del LS (dueño)
        const licenseDoc = {
          _id: new ObjectId(),
          id_employee: id, // el _id del LS que se va a LICENCIAS
          status: item.status || 1,
          ...item, // todos los campos del LS
          // Mantener los ID_CTRL_* del LS (no cambiar)
          ID_CTRL_ASIST: item.ID_CTRL_ASIST,
          ID_CTRL_TALON: item.ID_CTRL_TALON,
          ID_CTRL_NOM: item.ID_CTRL_NOM,
          ID_CTRL_CAP: item.ID_CTRL_CAP,
          ID_BITACORA: item.ID_BITACORA,
        };

        Object.keys(licenseDoc).forEach(
          (k) => licenseDoc[k] === undefined && delete licenseDoc[k]
        );

        licToInsert.push(licenseDoc);
        idsToDelete.push(id); // eliminar el LS de PLANTILLA
      }
    });

    if (licToInsert.length > 0) {
      await licCollection.insertMany(licToInsert, { ordered: false });
      console.log(
        `${licToInsert.length} dueños de plazas (TIPONOM 'LS') movidos a LICENCIAS.`
      );
    }

    if (idsToDelete.length > 0) {
      await collectionPlantilla.deleteMany({ _id: { $in: idsToDelete } });
      console.log(
        `${idsToDelete.length} dueños (LS) eliminados de PLANTILLA.`
      );
      // remove from jsonArray so later bulkWrite doesn't re-insert them
      const deletedIdsSet = new Set(idsToDelete.map((o) => o.toString()));
      jsonArray = jsonArray.filter((item, idx) => {
        const id = resultPlantilla.insertedIds[idx];
        return !deletedIdsSet.has(id.toString());
      });
    }

    if (licToInsert.length > 0) {
      await licCollection.updateMany(
        { TIPONOM: "LS" },
        { $set: { TIPONOM: "B" } }
      );
      console.log(
        `${licToInsert.length} documentos en LICENCIAS: TIPONOM actualizado de LS a B.`
      );
    }

    // Crear objetos bitácora después de insertar en PLANTILLA_2025
    const bitacoraArray = Object.values(resultPlantilla.insertedIds).map(
      (id, index) => {
        const empleado = originalJsonArray[index] || {};
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
    // --- Inicio: crear colección TALONES (un documento por empleado con array TALONES de 1 item para enero) ---
    console.log("Creating TALONES collection...");
    const collectionTalones = database.collection("TALONES");

    // Fecha de pago solo para la primera quincena de enero 2026 (quincena 1)
    const fechasPagoDiciembre = {
      1: "2026-01-15", // Primera quincena de enero (1-15)
    };

    const talonesArray = Object.values(resultPlantilla.insertedIds).map(
      (id) => {
        // Crear array de solo 1 talón (quincena 1 de enero)
        const talones = [];
        for (let quin = 1; quin <= 1; quin++) {
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
      `${talonesArray.length} documents inserted into TALONES collection (only January quincena 1).`
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

    // Agrupar por NUMPLA: si existe LS, usar solo ese; si no, usar el primero
    const plazasMap = new Map();
    for (const plaza of licenciaArray) {
      const numpla = plaza.NUMPLA;
      if (!numpla) continue;

      // Si no existe entrada para esta NUMPLA, o si la actual es LS (tienen prioridad)
      if (!plazasMap.has(numpla)) {
        plazasMap.set(numpla, plaza);
      } else if ((plaza.TIPONOM || "").toUpperCase() === "LS") {
        // Reemplazar si encontramos un LS (mayor prioridad)
        plazasMap.set(numpla, plaza);
      }
    }

    const plazasData = Array.from(plazasMap.values());
    await collectionPlazas.insertMany(plazasData);

    // Update plantilla before inserting into MongoDB
    const bulkOpsPlantillaUpdate = jsonArray
      .map((item) => ({
        updateOne: {
          filter: { ID_CTRL_ASIST: item.ID_CTRL_ASIST },
          update: { $set: item },
          upsert: true,
        },
      }))
      // in case jsonArray still contained deleted docs, ensure none of their IDs are processed
      .filter((op) => op.updateOne.filter.ID_CTRL_ASIST);

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

    // Crear incidencias vinculados a plantilla
    try {
      const collectionIncidencias = database.collection("INCIDENCIAS");
      await collectionIncidencias.deleteMany({});

      // Obtenemos el json de los permisos económicos ya creados
      const candidates = ["TRAMITES_EXISTENTES/incidencias.json"];
      let pathFound = null;
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          pathFound = p;
          break;
        }
      }

      if (!pathFound) {
        console.log("No se encontró el archivo de incidencias para crearlos.");
      } else {
        const raw = JSON.parse(fs.readFileSync(pathFound, "utf8"));
        const docsToInsert = [];
        const skipped = [];
        for (const item of raw) {
          const rfc = (item.RFC || "").toString().trim().toUpperCase();
          if (!rfc) {
            skipped.push({ original: item });
            continue;
          }

          const plantillaDoc = await collectionPlantilla.findOne({ RFC: rfc });
          if (!plantillaDoc) {
            skipped.push({ original: item });
            console.log(`El RFC: ${rfc} no se encontro en PLANTILLA, permiso económico no creado.`);
            continue;
          }

          const newDoc = { ...item };
          // Eliminar los id que vienene en el JSON
          delete newDoc._id;
          delete newDoc.ID_CTRL_ASIST;

          // Agregar el id del empleado y ID_CTRL_ASIST desde PLANTILLA
          newDoc.ID_CTRL_ASIST = plantillaDoc.ID_CTRL_ASIST || null;
          if (newDoc.ID_CTRL_ASIST && typeof newDoc.ID_CTRL_ASIST === "string") {
            try {
              newDoc.ID_CTRL_ASIST = new ObjectId(newDoc.ID_CTRL_ASIST);
            } catch (e) { }
          }

          newDoc.RFC = rfc;
          newDoc._id = new ObjectId();
          docsToInsert.push(newDoc);
        }

        if (docsToInsert.length > 0) {
          const res = await collectionIncidencias.insertMany(docsToInsert);
          console.log(`${Object.keys(res.insertedIds).length} incidencias creadas y vinculadas a PLANTILLA.`);
        } else {
          console.log("Incidencias no creadas, no se encontraron coincidencias con PLANTILLA.");
        }

        // Crear un JSON con los registros que no se pudieron crear por falta de coincidencia en PLANTILLA
        if (skipped.length > 0) {
          const outPath = "incidencias_no_creadas.json";
          try {
            fs.writeFileSync(outPath, JSON.stringify(skipped, null, 2), "utf8");
            console.log(`${skipped.length} incidencias sin crear escritas en ${outPath}`);
          } catch (werr) {
            console.error("Error al generar el archivo de incidencias no creadas:", werr);
          }
        }
      }
    } catch (errPerm) {
      console.error("Error creating incidencias:", errPerm);
    }

    // Crear permisos económicos vinculados a plantilla
    try {
      const collectionPermisosEconomicos = database.collection("PERMISOS_ECONOMICOS");
      await collectionPermisosEconomicos.deleteMany({});

      // Obtenemos el json de los permisos económicos ya creados
      const candidates = ["TRAMITES_EXISTENTES/permisos_economicos.json"];
      let pathFound = null;
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          pathFound = p;
          break;
        }
      }

      if (!pathFound) {
        console.log("No se encontró el archivo de permisos económicos para crearlos.");
      } else {
        const raw = JSON.parse(fs.readFileSync(pathFound, "utf8"));
        const docsToInsert = [];
        const skipped = [];
        for (const item of raw) {
          const rfc = (item.RFC || "").toString().trim().toUpperCase();
          if (!rfc) {
            skipped.push({ original: item });
            continue;
          }

          const plantillaDoc = await collectionPlantilla.findOne({ RFC: rfc });
          if (!plantillaDoc) {
            skipped.push({ original: item });
            console.log(`El RFC: ${rfc} no se encontro en PLANTILLA, permiso económico no creado.`);
            continue;
          }

          const newDoc = { ...item };
          // Eliminar los id que vienene en el JSON
          delete newDoc._id;
          delete newDoc.id_empoyee;
          delete newDoc.ID_CTRL_ASIST;

          // Agregar el id del empleado y ID_CTRL_ASIST desde PLANTILLA
          newDoc.id_empoyee = plantillaDoc._id;
          newDoc.ID_CTRL_ASIST = plantillaDoc.ID_CTRL_ASIST || null;
          if (newDoc.ID_CTRL_ASIST && typeof newDoc.ID_CTRL_ASIST === "string") {
            try {
              newDoc.ID_CTRL_ASIST = new ObjectId(newDoc.ID_CTRL_ASIST);
            } catch (e) { }
          }

          newDoc.RFC = rfc;
          newDoc._id = new ObjectId();
          docsToInsert.push(newDoc);
        }

        if (docsToInsert.length > 0) {
          const res = await collectionPermisosEconomicos.insertMany(docsToInsert);
          console.log(`${Object.keys(res.insertedIds).length} permisos económicos creados y vinculados a PLANTILLA.`);
        } else {
          console.log("Permisos económicos no creados, no se encontraron coincidencias con PLANTILLA.");
        }

        // Crear un JSON con los registros que no se pudieron crear por falta de coincidencia en PLANTILLA
        if (skipped.length > 0) {
          const outPath = "permisos_economicos_no_creados.json";
          try {
            fs.writeFileSync(outPath, JSON.stringify(skipped, null, 2), "utf8");
            console.log(`${skipped.length} permisos económicos sin crear escritos en ${outPath}`);
          } catch (werr) {
            console.error("Error al generar el archivo de permisos económicos no creados:", werr);
          }
        }
      }
    } catch (errPerm) {
      console.error("Error creating permisos económicos:", errPerm);
    }

    // Crear incapacidades vinculados a plantilla
    try {
      const collectionIncapacidades = database.collection("INCAPACIDADES");
      await collectionIncapacidades.deleteMany({});

      // Obtenemos el json de los permisos económicos ya creados
      const candidates = ["TRAMITES_EXISTENTES/incapacidades.json"];
      let pathFound = null;
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          pathFound = p;
          break;
        }
      }

      if (!pathFound) {
        console.log("No se encontró el archivo de incapacidades para crearlos.");
      } else {
        const raw = JSON.parse(fs.readFileSync(pathFound, "utf8"));
        const docsToInsert = [];
        const skipped = [];
        for (const item of raw) {
          const rfc = (item.RFC || "").toString().trim().toUpperCase();
          if (!rfc) {
            skipped.push({ original: item });
            continue;
          }

          const plantillaDoc = await collectionPlantilla.findOne({ RFC: rfc });
          if (!plantillaDoc) {
            skipped.push({ original: item });
            console.log(`El RFC: ${rfc} no se encontro en PLANTILLA, incapacidad no creada.`);
            continue;
          }

          const newDoc = { ...item };
          delete newDoc._id;
          delete newDoc.id_empoyee;
          delete newDoc.ID_CTRL_ASIST;
          delete newDoc.RFC;

          newDoc.id_empoyee = plantillaDoc._id;
          newDoc.ID_CTRL_ASIST = plantillaDoc.ID_CTRL_ASIST || null;
          if (newDoc.ID_CTRL_ASIST && typeof newDoc.ID_CTRL_ASIST === "string") {
            try {
              newDoc.ID_CTRL_ASIST = new ObjectId(newDoc.ID_CTRL_ASIST);
            } catch (e) { }
          }

          newDoc._id = new ObjectId();
          docsToInsert.push(newDoc);
        }

        if (docsToInsert.length > 0) {
          const res = await collectionIncapacidades.insertMany(docsToInsert);
          console.log(`${Object.keys(res.insertedIds).length} incapacidades creadas y vinculadas a PLANTILLA.`);
        } else {
          console.log("Incapacidades no creadas, no se encontraron coincidencias con PLANTILLA.");
        }

        // Crear un JSON con los registros que no se pudieron crear por falta de coincidencia en PLANTILLA
        if (skipped.length > 0) {
          const outPath = "incapacidades_no_creadas.json";
          try {
            fs.writeFileSync(outPath, JSON.stringify(skipped, null, 2), "utf8");
            console.log(`${skipped.length} incapacidades sin crear escritos en ${outPath}`);
          } catch (werr) {
            console.error("Error al generar el archivo de incapacidades no creadas:", werr);
          }
        }
      }
    } catch (errPerm) {
      console.error("Error creating incapacidades:", errPerm);
    }

    // Crear justificaciones vinculados a plantilla
    try {
      const collectionJustificaciones = database.collection("JUSTIFICACIONES");
      await collectionJustificaciones.deleteMany({});

      const candidates = ["TRAMITES_EXISTENTES/justificaciones.json"];
      let pathFound = null;
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          pathFound = p;
          break;
        }
      }

      if (!pathFound) {
        console.log("No se encontró el archivo de justificaciones para crearlos.");
      } else {
        const raw = JSON.parse(fs.readFileSync(pathFound, "utf8"));
        const docsToInsert = [];
        const skipped = [];
        for (const item of raw) {
          const rfc = (item.RFC || "").toString().trim().toUpperCase();
          if (!rfc) {
            skipped.push({ original: item });
            continue;
          }

          const plantillaDoc = await collectionPlantilla.findOne({ RFC: rfc });
          if (!plantillaDoc) {
            skipped.push({ original: item });
            console.log(`El RFC: ${rfc} no se encontro en PLANTILLA, justificacion no creada.`);
            continue;
          }

          const newDoc = { ...item };
          delete newDoc._id;
          delete newDoc.id_empoyee;
          delete newDoc.ID_CTRL_ASIST;
          delete newDoc.RFC;

          newDoc.id_empoyee = plantillaDoc._id;
          newDoc.ID_CTRL_ASIST = plantillaDoc.ID_CTRL_ASIST || null;
          if (newDoc.ID_CTRL_ASIST && typeof newDoc.ID_CTRL_ASIST === "string") {
            try {
              newDoc.ID_CTRL_ASIST = new ObjectId(newDoc.ID_CTRL_ASIST);
            } catch (e) { }
          }

          newDoc._id = new ObjectId();
          docsToInsert.push(newDoc);
        }

        if (docsToInsert.length > 0) {
          const res = await collectionJustificaciones.insertMany(docsToInsert);
          console.log(`${Object.keys(res.insertedIds).length} justificaciones creadas y vinculadas a PLANTILLA.`);
        } else {
          console.log("Justificaciones no creadas, no se encontraron coincidencias con PLANTILLA.");
        }

        if (skipped.length > 0) {
          const outPath = "justificaciones_no_creadas.json";
          try {
            fs.writeFileSync(outPath, JSON.stringify(skipped, null, 2), "utf8");
            console.log(`${skipped.length} justificaciones sin crear escritos en ${outPath}`);
          } catch (werr) {
            console.error("Error al generar el archivo de justificaciones no creadas:", werr);
          }
        }
      }
    } catch (errPerm) {
      console.error("Error creating justificaciones:", errPerm);
    }

    // Crear permisos extraordinarios vinculados a plantilla
    try {
      const collectionPermisosExt = database.collection("PERMISOS_EXT");
      await collectionPermisosExt.deleteMany({});

      const candidates = ["TRAMITES_EXISTENTES/permisos_ext.json"];
      let pathFound = null;
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          pathFound = p;
          break;
        }
      }

      if (!pathFound) {
        console.log("No se encontró el archivo de permisos extraordinarios para crearlos.");
      } else {
        const raw = JSON.parse(fs.readFileSync(pathFound, "utf8"));
        const docsToInsert = [];
        const skipped = [];
        for (const item of raw) {
          const rfc = (item.RFC || "").toString().trim().toUpperCase();
          if (!rfc) {
            skipped.push({ original: item });
            continue;
          }

          const plantillaDoc = await collectionPlantilla.findOne({ RFC: rfc });
          if (!plantillaDoc) {
            skipped.push({ original: item });
            console.log(`El RFC: ${rfc} no se encontro en PLANTILLA, permiso extraordinario no creado.`);
            continue;
          }

          const newDoc = { ...item };
          delete newDoc._id;
          delete newDoc.id_empoyee;
          delete newDoc.ID_CTRL_ASIST;
          delete newDoc.RFC;

          newDoc.id_empoyee = plantillaDoc._id;
          newDoc.ID_CTRL_ASIST = plantillaDoc.ID_CTRL_ASIST || null;
          if (newDoc.ID_CTRL_ASIST && typeof newDoc.ID_CTRL_ASIST === "string") {
            try {
              newDoc.ID_CTRL_ASIST = new ObjectId(newDoc.ID_CTRL_ASIST);
            } catch (e) { }
          }

          newDoc._id = new ObjectId();
          docsToInsert.push(newDoc);
        }

        if (docsToInsert.length > 0) {
          const res = await collectionPermisosExt.insertMany(docsToInsert);
          console.log(`${Object.keys(res.insertedIds).length} permisos extraordinarios creadas y vinculadas a PLANTILLA.`);
        } else {
          console.log("Permisos extraordinarios no creadas, no se encontraron coincidencias con PLANTILLA.");
        }

        if (skipped.length > 0) {
          const outPath = "permisos_ext_no_creados.json";
          try {
            fs.writeFileSync(outPath, JSON.stringify(skipped, null, 2), "utf8");
            console.log(`${skipped.length} permisos extraordinarios sin crear escritos en ${outPath}`);
          } catch (werr) {
            console.error("Error al generar el archivo de permisos extraordinarios no creados:", werr);
          }
        }
      }
    } catch (errPerm) {
      console.error("Error creating permisos extraordinarios:", errPerm);
    }

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
