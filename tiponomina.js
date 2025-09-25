const { MongoClient } = require("mongodb");
const mysql = require("mysql");
const fs = require("fs");

// Configuración de la conexión a MongoDB
const uri = "mongodb://localhost:27017"; // Cambia esto según tu configuración
const dbName = "sirhTest";
const collectionName = "PLANTILLA";
const collectionLicencias = "LICENCIAS"; // Cambia esto por el nombre de tu colección

// Configuración de la conexión a MySQL
const mysqlConnection = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "",
  database: "sirh",
});

// Conectar a la base de datos MySQL
mysqlConnection.connect((err) => {
  if (err) {
    console.error("Error al conectar a MySQL:", err);
    return;
  }
  console.log("Conexión a MySQL establecida correctamente.");
});

async function procesarPlantillatipoNOM() {
  const client = new MongoClient(uri);
  const noCoincidentes = []; // Array para almacenar los registros sin coincidencia

  try {
    // Conexión a la base de datos
    await client.connect();
    const db = client.db(dbName);
    const collection = db.collection(collectionName);

    // Consulta a la colección PLANTILLA
    const registros = await collection.find({}).toArray();

    // Filtrar registros con TIPONOM igual a 'LS'
    const registrosLicencias = registros.filter(
      (registro) => registro.TIPONOM === "LS"
    );
    registrosLicencias.forEach((licencia) => (licencia.status = 1));
    if (registrosLicencias.length > 0) {
      const collectionLic = db.collection(collectionLicencias);
      await collectionLic.insertMany(registrosLicencias);
      console.log(
        "Registros con TIPONOM 'LS' insertados en la colección LICENCIAS."
      );
    }

    // Consulta a la tabla categorias_catalogo en MySQL
    const categorias = await new Promise((resolve, reject) => {
      mysqlConnection.query(
        "SELECT * FROM categorias_catalogo",
        (err, results) => {
          if (err) {
            return reject(err);
          }
          resolve(results);
        }
      );
    });

    console.log("Categorías obtenidas correctamente.");

    // Procesar cada registro
    for (const registro of registros) {
      const categoriaCoincidente = categorias.find(
        (categoria) => categoria.CLAVE_CATEGORIA === registro.CLAVECAT
      );

      if (categoriaCoincidente) {
        // Actualizar las propiedades del registro
        registro.TIPONOM = categoriaCoincidente.T_NOMINA;
        registro.NIVEL = categoriaCoincidente.NIVEL;
        registro.NOMCATE = categoriaCoincidente.DESCRIPCION;

        // Actualizar el registro en la base de datos
        await collection.updateOne({ _id: registro._id }, { $set: registro });
        console.log(
          `Registro con CLAVECAT ${registro.CLAVECAT} actualizado correctamente.`
        );
      } else {
        // Agregar el registro al array de no coincidentes
        noCoincidentes.push(registro);
      }
    }

    // Escribir los registros no coincidentes en un archivo JSON
    if (noCoincidentes.length > 0) {
      fs.writeFileSync(
        "noCoincidentes.json",
        JSON.stringify(noCoincidentes, null, 2),
        "utf-8"
      );
      console.log(
        "Registros no coincidentes guardados en noCoincidentes.json."
      );
    }
  } catch (error) {
    console.error("Error al procesar la plantilla:", error);
  } finally {
    // Cerrar la conexión
    await client.close();
    mysqlConnection.end(); // Finalizar la conexión MySQL
    console.log("Proceso completado y conexiones cerradas.");
  }
}

async function actualizarTiponomEnPlazas() {
  await new Promise((resolve) => {
    setTimeout(() => {
      console.log("Iniciando actualización de TIPONOM en PLAZAS...");
      resolve();
    }, 5000);
  });
  const client = new MongoClient(uri);
  const collectionPlazas = "PLAZAS"; // Nombre de la colección PLAZAS

  try {
    // Conexión a la base de datos
    await client.connect();
    const db = client.db(dbName);
    const collectionPlantilla = db.collection(collectionName);
    const collectionPlazas = db.collection("PLAZAS"); // Cambia esto por el nombre de tu colección PLAZAS

    // Obtener todos los registros de la colección PLANTILLA
    const registrosPlantilla = await collectionPlantilla.find({}).toArray();

    // Procesar cada registro de la colección PLANTILLA
    for (const registroPlantilla of registrosPlantilla) {
      // Buscar coincidencias en la colección PLAZAS
      const resultadoPlaza = await collectionPlazas.findOne({
        NUMPLA: registroPlantilla.NUMPLA,
      });

      if (resultadoPlaza) {
        // Actualizar el valor de TIPONOM en la colección PLAZAS
        await collectionPlazas.updateOne(
          { _id: resultadoPlaza._id },
          { $set: { TIPONOM: registroPlantilla.TIPONOM } }
        );
        console.log(
          `TIPONOM actualizado en PLAZAS para NUMPLA ${registroPlantilla.NUMPLA}.`
        );
      }
    }
  } catch (error) {
    console.error("Error al actualizar TIPONOM en PLAZAS:", error);
  } finally {
    // Cerrar la conexión
    await client.close();
    console.log("Proceso de actualización completado y conexión cerrada.");
  }
}
module.exports = { procesarPlantillatipoNOM, actualizarTiponomEnPlazas };
