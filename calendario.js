const { MongoClient } = require("mongodb");
const moment = require("moment");
require("moment/locale/es-mx"); // Importa el locale español de México
moment.locale("es-mx"); // Establece el locale a español de México

const uri = "mongodb://localhost:27017";

// Array de días inhábiles adicionales con festividades
const diasInhabiles = [
  { fecha: "01-01-2026", festividad: "AÑO NUEVO" },
  { fecha: "03-02-2026", festividad: "DÍA DE LA CONSTITUCIÓN" },
  { fecha: "17-03-2026", festividad: "NATALICIO DE BENITO JUÁREZ" },
  { fecha: "01-05-2026", festividad: "DÍA DEL TRABAJO" },
  { fecha: "16-09-2026", festividad: "DÍA DE LA INDEPENDENCIA" },
  { fecha: "18-11-2026", festividad: "REVOLUCIÓN MEXICANA" },
  { fecha: "25-12-2026", festividad: "NAVIDAD" },
  { fecha: "02-04-2026", festividad: "SEMANA SANTA" },
  { fecha: "03-04-2026", festividad: "SEMANA SANTA" },
  { fecha: "05-05-2026", festividad: "BATALLA DE PUEBLA" },
  { fecha: "10-05-2026", festividad: "DÍA DE LA MADRE" },
  { fecha: "18-07-2026", festividad: "ANIVERSARIO LUCTUOSO DE BENITO JUÁREZ" },
  { fecha: "20-07-2026", festividad: "PRIMER LUNES DEL CERRO" },
  { fecha: "27-07-2026", festividad: "SEGUNDO LUNES DEL CERRO" },
  { fecha: "21-10-2026", festividad: "DÍA DEL EMPLEADO OAXACA" },
  { fecha: "01-11-2026", festividad: "DÍA DE MUERTOS" },
  { fecha: "02-11-2026", festividad: "DÍA DE MUERTOS" },
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
    const fecha = moment("2026-01-01").add(i, "days");
    const diaDeLaSemana = fecha.format("dddd").toUpperCase();
    const esInhabil = diasInhabiles.some(
      (d) => d.fecha === fecha.format("DD-MM-YYYY")
    );
    const esFinDeSemana = ["SÁBADO", "DOMINGO"].includes(diaDeLaSemana);

    // Calcula la quincena correctamente basándose en el día del mes y el mes del año
    const diaDelMes = fecha.date();
    const mes = fecha.month() + 1; // moment().month() retorna 0-11, por eso sumamos 1
    // Primera quincena: días 1-15, Segunda quincena: días 16-fin de mes
    // La quincena del año se calcula: (mes-1)*2 + (si es día 16 o más, entonces 2, sino 1)
    const quincena = (mes - 1) * 2 + (diaDelMes <= 15 ? 1 : 2);

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
    const databaseName = "SIRH2026";
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
