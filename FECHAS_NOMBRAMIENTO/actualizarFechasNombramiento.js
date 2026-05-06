const { MongoClient } = require('mongodb');
const XLSX = require('xlsx');

// Configuración
const mongoUri = 'mongodb://localhost:27017'; // Ajusta si es remoto
const dbName = 'SIRH2026';
const excelInput = 'FECHAS_NOMBRAMIENTO/Personal_fecha_nombramiento.xlsx'; // Nombre del Excel de entrada
const excelOutput = 'FECHAS_NOMBRAMIENTO/rfcs_no_encontrados.xlsx'; // Nombre del Excel de salida
const rfcColumn = 'RFC'; // Columna del RFC en el Excel
const nombreColumn = 'NOMBRE'; // Columna del NOMBRE en el Excel (si es necesario)
const numplaColumn = 'NUMPLA'; // Columna del NUMPLA en el Excel (si es necesario)
const fechaColumn = 'INGRESO'; // Columna de la fecha en el Excel (cambia si es 'FECHA_NOMBRAMIENTO')

async function main() {
    const client = new MongoClient(mongoUri);

    try {
        // Conectar a MongoDB
        await client.connect();
        console.log('Conectado a MongoDB');
        const db = client.db(dbName);
        const plantilla = db.collection('PLANTILLA');
        const licencias = db.collection('LICENCIAS');

        // Leer el Excel de entrada
        const workbook = XLSX.readFile(excelInput);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        const excelRfcs = new Set(
            data
                .map(row => row[rfcColumn])
                .filter(rfc => rfc !== undefined && rfc !== null)
                .map(rfc => String(rfc).trim())
                .filter(rfc => rfc !== '')
        );

        const rfcsNoEncontrados = [];

        // Procesar cada fila
        for (const row of data) {
            const rfc = row[rfcColumn];
            const nombre = row[nombreColumn];
            const numpla = row[numplaColumn];
            const fechaStr = row[fechaColumn];

            if (!rfc || !fechaStr) {
                console.warn(`Fila incompleta: RFC=${rfc}, Fecha=${fechaStr}`);
                continue;
            }

            function parseFecha(valor) {
                // Caso 1: Excel como número (serial)
                if (typeof valor === 'number') {
                    const fecha = XLSX.SSF.parse_date_code(valor);
                    return new Date(fecha.y, fecha.m - 1, fecha.d);
                }

                // Caso 2: ya es Date
                if (valor instanceof Date) {
                    return valor;
                }

                // Caso 3: string tipo "DD/MM/YYYY"
                if (typeof valor === 'string') {
                    const [dia, mes, anio] = valor.split('/').map(Number);
                    return new Date(anio, mes - 1, dia);
                }

                // Caso inválido
                return null;
            }

            // Convertir fecha a objeto Date (asumiendo formato DD/MM/YYYY o similar; ajusta si es ISO)
            const fecha = parseFecha(fechaStr);

            if (!fecha || isNaN(fecha.getTime())) {
                console.warn(`Fecha inválida para RFC ${rfc}:`, fechaStr);
                continue;
            }

            // Buscar documento por RFC
            const query = { RFC: rfc };
            const update = { $set: { FECHA_NOMBRAMIENTO: fecha } };
            let result = await plantilla.updateOne(query, update, { upsert: false }); // No upsert, solo actualizar si existe

            if (result.matchedCount === 0) {
                result = await licencias.updateOne(query, update);
                if (result.matchedCount === 0) {
                    // No encontrado, agregar a lista
                    rfcsNoEncontrados.push({ RFC: rfc, NOMBRE: nombre, NUMPLA: numpla, INGRESO: fecha });
                    console.log(`RFC no encontrado: ${rfc}`);
                } else {
                    console.log(`Actualizado en LICENCIAS: ${rfc}`);
                }
            } else {
                console.log(`Actualizado RFC: ${rfc}`);
            }
        }

        // Establecer FECHA_NOMBRAMIENTO a null para registros no incluidos en el Excel
        if (excelRfcs.size > 0) {
            const rfcsArray = Array.from(excelRfcs);
            const updateNull = { $set: { FECHA_NOMBRAMIENTO: null } };

            const plantillaResult = await plantilla.updateMany(
                { RFC: { $nin: rfcsArray } },
                updateNull
            );
            console.log(`PLANTILLA: ${plantillaResult.modifiedCount} documentos actualizados a null`);

            const licenciasResult = await licencias.updateMany(
                { RFC: { $nin: rfcsArray } },
                updateNull
            );
            console.log(`LICENCIAS: ${licenciasResult.modifiedCount} documentos actualizados a null`);
        }

        // Crear Excel con RFCs no encontrados
        if (rfcsNoEncontrados.length > 0) {
            const ws = XLSX.utils.json_to_sheet(rfcsNoEncontrados);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'NoEncontrados');
            XLSX.writeFile(wb, excelOutput);
            console.log(`Excel creado: ${excelOutput} con ${rfcsNoEncontrados.length} RFCs no encontrados`);
        } else {
            console.log('Todos los RFCs fueron encontrados y actualizados.');
        }

    } catch (error) {
        console.error('Error:', error);
    } finally {
        await client.close();
        console.log('Conexión cerrada');
    }
}

main();