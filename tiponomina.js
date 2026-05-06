const { MongoClient } = require("mongodb");
const fs = require("fs");

// Configuración de la conexión a MongoDB
const uri = "mongodb://localhost:27017";
const dbName = "SIRH2026";
const projectField = "PROYECTO";
const projectList = [
  "114004148010000010@",
  "1140041480100000101",
  "1140041480100000102",
  "1140041480100000103",
  "1140041480100000104",
  "1140041480100000105",
  "1140041480100000106",
  "1140041480100000107",
  "1140041480100000108",
  "1140041480100000109",
  "114004148010000010A",
  "114004148010000010B",
  "114004148010000010C",
  "114004148010000010D",
  "114004148010000010E",
  "114004148010000010F",
  "114004148010000010G",
  "114004148010000010H",
  "114004148010000010I",
  "114004148010000010J",
  "114004148010000010K",
  "114004148010000010L",
  "114004148010000010M",
  "114004148010000010N",
  "114004148010000010O",
  "114004148010000010P",
  "114004148010000010Q",
  "114004148010000010R",
  "114004148010000010S",
  "114004148010000010T",
  "114004148010000010U",
  "114004148010000010W",
  "114004148010000010X",
  "114004148010000010Z",
];

// Helper que procesa TIPONOM en cualquier colección
async function procesarTipoNOMEn(collectionName) {
  const client = new MongoClient(uri);
  const noCoincidentes = [];

  try {
    await client.connect();
    const db = client.db(dbName);
    const col = db.collection(collectionName);

    const registros = await col.find({}).toArray();

    for (const registro of registros) {
      const proyecto = registro[projectField];
      const tipOriginal = (registro.TIPONOM || "").toUpperCase();

      if (tipOriginal === "B") {
        if (!proyecto) {
          noCoincidentes.push(registro);
          continue;
        }
        const nuevoTip = projectList.includes(proyecto) ? "F51" : "M51";
        if (nuevoTip !== registro.TIPONOM) {
          await col.updateOne(
            { _id: registro._id },
            { $set: { TIPONOM: nuevoTip } }
          );
          console.log(
            `${collectionName} NUMPLA ${registro.NUMPLA || registro._id}: TIPONOM '${registro.TIPONOM}' -> '${nuevoTip}'.`
          );
        }
        continue;
      }

      if (tipOriginal === "CC") {
        if (!proyecto) {
          noCoincidentes.push(registro);
          continue;
        }
        const nuevoTip = projectList.includes(proyecto) ? "FCT" : "CCT";
        if (nuevoTip !== registro.TIPONOM) {
          await col.updateOne(
            { _id: registro._id },
            { $set: { TIPONOM: nuevoTip } }
          );
          console.log(
            `${collectionName} NUMPLA ${registro.NUMPLA || registro._id}: TIPONOM '${registro.TIPONOM}' -> '${nuevoTip}'.`
          );
        }
        continue;
      }

      if (tipOriginal === "CN") {
        if (!proyecto) {
          noCoincidentes.push(registro);
          continue;
        }
        const nuevoTip = projectList.includes(proyecto) ? "FCO" : "511";
        if (nuevoTip !== registro.TIPONOM) {
          await col.updateOne(
            { _id: registro._id },
            { $set: { TIPONOM: nuevoTip } }
          );
          console.log(
            `${collectionName} NUMPLA ${registro.NUMPLA || registro._id}: TIPONOM '${registro.TIPONOM}' -> '${nuevoTip}'.`
          );
        }
        continue;
      }

      if (tipOriginal === "MM") {
        if (!proyecto) {
          noCoincidentes.push(registro);
          continue;
        }
        const nuevoTip = projectList.includes(proyecto) ? "FMM" : "MMS";
        if (nuevoTip !== registro.TIPONOM) {
          await col.updateOne(
            { _id: registro._id },
            { $set: { TIPONOM: nuevoTip } }
          );
          console.log(
            `${collectionName} NUMPLA ${registro.NUMPLA || registro._id}: TIPONOM '${registro.TIPONOM}' -> '${nuevoTip}'.`
          );
        }
        continue;
      }

      if (tipOriginal && tipOriginal !== "LS") {
        noCoincidentes.push(registro);
      }
    }

    if (noCoincidentes.length > 0) {
      fs.writeFileSync(
        "noCoincidentes.json",
        JSON.stringify(noCoincidentes, null, 2),
        "utf-8"
      );
      console.log("Registros no coincidentes guardados en noCoincidentes.json.");
    }
  } catch (error) {
    console.error(`Error al procesar ${collectionName}:`, error);
  } finally {
    await client.close();
  }
}

// Función pública que ejecuta el helper sobre PLANTILLA y LICENCIAS
async function procesarPlantillatipoNOM() {
  await procesarTipoNOMEn("PLANTILLA");
  await procesarTipoNOMEn("LICENCIAS");
}

// Actualizar TIPONOM en PLAZAS basándose en PLANTILLA
async function actualizarTiponomEnPlazas() {
  await new Promise((resolve) => {
    setTimeout(() => {
      console.log("Iniciando actualización de TIPONOM en PLAZAS...");
      resolve();
    }, 5000);
  });

  const client = new MongoClient(uri);

  try {
    await client.connect();
    const db = client.db(dbName);
    const collectionPlantilla = db.collection("PLANTILLA");
    const collectionPlazas = db.collection("PLAZAS");

    const registrosPlantilla = await collectionPlantilla.find({}).toArray();

    for (const registroPlantilla of registrosPlantilla) {
      const resultadoPlaza = await collectionPlazas.findOne({
        NUMPLA: registroPlantilla.NUMPLA,
      });

      if (resultadoPlaza) {
        await collectionPlazas.updateOne(
          { _id: resultadoPlaza._id },
          { $set: { TIPONOM: registroPlantilla.TIPONOM } }
        );
        console.log(
          `TIPONOM actualizado en PLAZAS para NUMPLA ${registroPlantilla.NUMPLA}: ${registroPlantilla.TIPONOM}`
        );
      }
    }
  } catch (error) {
    console.error("Error al actualizar TIPONOM en PLAZAS:", error);
  } finally {
    await client.close();
    console.log("Proceso de actualización completado y conexión cerrada.");
  }
}

module.exports = { procesarPlantillatipoNOM, actualizarTiponomEnPlazas };