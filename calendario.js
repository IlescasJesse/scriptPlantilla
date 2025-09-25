const { MongoClient } = require("mongodb");
const moment = require("moment");
require("moment/locale/es-mx"); // Importa el locale español de México
moment.locale("es-mx"); // Establece el locale a español de México

const uri = "mongodb://localhost:27017";

// Array de días inhábiles adicionales con festividades
const diasInhabiles = [
  { fecha: "01-01-2025", festividad: "AÑO NUEVO" },
  { fecha: "03-02-2025", festividad: "DÍA DE LA CONSTITUCIÓN" },
  { fecha: "17-03-2025", festividad: "NATALICIO DE BENITO JUÁREZ" },
  { fecha: "01-05-2025", festividad: "DÍA DEL TRABAJO" },
  { fecha: "16-09-2025", festividad: "DÍA DE LA INDEPENDENCIA" },
  { fecha: "18-11-2025", festividad: "REVOLUCIÓN MEXICANA" },
  { fecha: "25-12-2025", festividad: "NAVIDAD" },
  { fecha: "17-04-2025", festividad: "SEMANA SANTA" },
  { fecha: "18-04-2025", festividad: "SEMANA SANTA" },
  { fecha: "05-05-2025", festividad: "BATALLA DE PUEBLA" },
  { fecha: "10-05-2025", festividad: "DÍA DE LA MADRE" },
  { fecha: "18-07-2025", festividad: "ANIVERSARIO LUCTUOSO DE BENITO JUÁREZ" },
  { fecha: "21-07-2025", festividad: "PRIMER LUNES DEL CERRO" },
  { fecha: "28-07-2025", festividad: "SEGUNDO LUNES DEL CERRO" },
  { fecha: "21-10-2025", festividad: "DÍA DEL EMPLEADO OAXACA" },
  { fecha: "01-11-2025", festividad: "DÍA DE MUERTOS" },
  { fecha: "02-11-2025", festividad: "DÍA DE MUERTOS" },
];

async function eliminarColeccionSiExiste(client, databaseName, collectionName) {
  const database = client.db(databaseName);
  const collectionNames = await database
    .listCollections({ name: collectionName })
    .toArray();

  if (collectionNames.length > 0) {
    await database.collection(collectionName).drop();
    console.log(`Colección ${collectionName} eliminada exitosamente.`);
  }
}

async function insertarDatos(client, databaseName, collectionName) {
  const database = client.db(databaseName);
  const collection = database.collection(collectionName);

  const dias = [];
  for (let i = 0; i < 365; i++) {
    const fecha = moment("2025-01-01").add(i, "days");
    const diaDeLaSemana = fecha.format("dddd").toUpperCase();
    const esInhabil = diasInhabiles.some(
      (d) => d.fecha === fecha.format("DD-MM-YYYY")
    );
    const esFinDeSemana = ["SÁBADO", "DOMINGO"].includes(diaDeLaSemana);
    const quincena = Math.floor(i / 15) + 1;
    const festividad =
      diasInhabiles.find((d) => d.fecha === fecha.format("DD-MM-YYYY"))
        ?.festividad || (esFinDeSemana ? diaDeLaSemana : null);
    const dia = {
      FECHA: fecha.format("DD-MM-YYYY"),
      DIA: diaDeLaSemana,
      HABIL: {
        BASE: !esInhabil && !esFinDeSemana,
        CONTRATO: !esInhabil && !esFinDeSemana,
      },
      QUIN: quincena,
      MOTIVO: festividad,
    };
    dias.push(dia);
  }

  // Inserta los documentos en la colección
  await collection.insertMany(dias);
  console.log("Días insertados exitosamente en CALENDARIO");
}

async function main() {
  const client = new MongoClient(uri, {});

  try {
    await client.connect();
    const databaseName = "sirhTest";
    const collectionName = "CALENDARIO";

    // Elimina la colección si existe
    await eliminarColeccionSiExiste(client, databaseName, collectionName);

    // Inserta los datos
    await insertarDatos(client, databaseName, collectionName);
  } finally {
    await client.close();
  }
}

main().catch(console.dir);
