const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { MongoClient } = require("mongodb");

// Rutas de los archivos
const asignacionesPath = path.join(__dirname, "ASIGNACIONES_LABORALES.xlsx");
const comisionadosPath = path.join(__dirname, "COMISIONADOS_SINDICALES.xlsx");

// Leer los archivos Excel
const asignacionesWorkbook = xlsx.readFile(asignacionesPath);
const comisionadosWorkbook = xlsx.readFile(comisionadosPath);

// Obtener las hojas de trabajo
const asignacionesSheet = asignacionesWorkbook.Sheets["asignaciones"];
const comisionadosSheet = comisionadosWorkbook.Sheets["comisionados"];

// Convertir las hojas a JSON
const asignacionesData = xlsx.utils.sheet_to_json(asignacionesSheet);
const comisionadosData = xlsx.utils.sheet_to_json(comisionadosSheet);

// Configuración de MongoDB
const uri = "mongodb://admin:1234@localhost:27017/";
const client = new MongoClient(uri);

const actualizarPlantillaDesdeMongo = async () => {
  try {
    console.log("Conectando a MongoDB...");
    await client.connect();
    const database = client.db("sirhTest");
    const collectionPlantilla = database.collection("PLANTILLA");

    console.log("Obteniendo registros de la colección PLANTILLA...");
    const plantilla = await collectionPlantilla.find().toArray();

    const operaciones = plantilla.map((empleado) => {
      const nombreCompletoPlantilla =
        `${empleado.APE_PAT} ${empleado.APE_MAT} ${empleado.NOMBRES}`.trim();

      // Buscar en ASIGNACIONES_LABORALES
      const asignacion = asignacionesData.find((asig) => {
        const nombreCompletoAsignacion = asig.NOMBRE.trim();
        return nombreCompletoPlantilla === nombreCompletoAsignacion;
      });

      if (asignacion) {
        return {
          updateOne: {
            filter: { _id: empleado._id },
            update: {
              $set: {
                STATUS_EMPLEADO: {
                  STATUS: "ASIG_LAB",
                  DESDE: null,
                  HASTA: null,
                  LUGAR_COMISIONADO: asignacion.AREA || null,
                  OBSERVACIONES: null,
                  PRYECTO: null,
                },
              },
            },
          },
        };
      }

      // Buscar en COMISIONADOS_SINDICALES
      const comisionado = comisionadosData.find((com) => {
        const nombreCompletoComisionado = com.NOMBRE.trim();
        return nombreCompletoPlantilla === nombreCompletoComisionado;
      });

      if (comisionado) {
        return {
          updateOne: {
            filter: { _id: empleado._id },
            update: {
              $set: {
                STATUS_EMPLEADO: {
                  STATUS: "COM_SDCL",
                  DESDE: null,
                  HASTA: null,
                  LUGAR_COMISIONADO: "STPEIDCEO",
                  OBSERVACIONES: comisionado.OBSERVACIONES || null,
                },
              },
            },
          },
        };
      }

      // Si no coincide con ninguno, asignar STATUS_EMPLEADO como null
      return {
        updateOne: {
          filter: { _id: empleado._id },
          update: {
            $set: {
              STATUS_EMPLEADO: null,
            },
          },
        },
      };
    });
    operaciones.forEach((operacion) => {
      if (operacion.updateOne.update.$set.STATUS_EMPLEADO) {
        const statusEmpleado = operacion.updateOne.update.$set.STATUS_EMPLEADO;
        const nombreCompleto = plantilla.find(
          (empleado) => empleado._id === operacion.updateOne.filter._id
        );
        if (nombreCompleto) {
          // console.log(
          //   `Registro actualizado: ${nombreCompleto.APE_PAT} ${nombreCompleto.APE_MAT} ${nombreCompleto.NOMBRES} - STATUS: ${statusEmpleado.STATUS}`
          // );
        }
      }
    });

    console.log("Actualizando registros en la colección PLANTILLA...");
    if (operaciones.length > 0) {
      await collectionPlantilla.bulkWrite(operaciones);
    }

    console.log("Actualización de PLANTILLA completada.");
  } catch (error) {
    console.error("Error al actualizar la plantilla:", error);
  } finally {
    console.log("Cerrando conexión con MongoDB...");
    await client.close();
    console.log("Conexión cerrada.");
  }
};

// Exportar la función
module.exports = { actualizarPlantillaDesdeMongo };
