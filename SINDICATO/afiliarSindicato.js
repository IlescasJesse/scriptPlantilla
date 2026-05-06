const { MongoClient } = require('mongodb');
const XLSX = require('xlsx');

// Configuración
const mongoUri = 'mongodb://localhost:27017'; // Ajusta si es remoto
const dbName = 'SIRH2026';
const excelInput = 'SINDICATO/Personal_afiliado.xlsx'; // Nombre del Excel de entrada
const excelOutput = 'SINDICATO/rfcs_no_encontrados.xlsx'; // Nombre del Excel de salida
const delegadosExcel = 'SINDICATO/delegados y agremiados.xlsx';
const rfcColumn = 'RFC'; // Columna del RFC en el Excel
const nombreColumn = 'NOMBRE'; // Columna del NOMBRE en el Excel (si es necesario)
const numplaColumn = 'NUMPLA'; // Columna del NUMPLA en el Excel (si es necesario)

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

        // Nuevo: Leer el Excel de delegados y agremiados
        const delegadosWorkbook = XLSX.readFile(delegadosExcel);
        const delegadosSheetName = delegadosWorkbook.SheetNames[0];
        const delegadosWorksheet = delegadosWorkbook.Sheets[delegadosSheetName];
        const delegadosData = XLSX.utils.sheet_to_json(delegadosWorksheet);

        const rfcsNoEncontrados = [];

        // Procesar cada fila
        for (const row of data) {
            const rfc = row[rfcColumn];
            const nombre = row[nombreColumn];
            const numpla = row[numplaColumn];

            if (!rfc) {
                console.warn(`Fila incompleta: RFC=${rfc}`);
                continue;
            }

            // Buscar documento por RFC
            const query = { RFC: rfc };

            let doc = await plantilla.findOne(query);
            let collectionName = 'PLANTILLA';
            if (!doc) {
                doc = await licencias.findOne(query);
                collectionName = 'LICENCIAS';
            }

            let delegacion = '';
            let delegado = '';

            if (doc) {
                // 1. Nombre oficial desde Mongo
                const fullNameDB = `${doc.APE_PAT || ''} ${doc.APE_MAT || ''} ${doc.NOMBRES || ''}`.trim();

                // 2. Buscar en Excel de delegados por NOMBRE (no RFC)
                const delegadosRow = delegadosData.find(d => {
                    const fullNameExcel = (d['APELLIDOS Y NOMBRES'] || '').trim();
                    return fullNameDB === fullNameExcel;
                });

                if (delegadosRow) {
                    delegacion = delegadosRow.DELEGACION || '';
                    delegado = delegadosRow.DELEGADOS || '';
                } else {
                    console.log(`Nombre no encontrado en Excel de delegados: "${fullNameDB}"`);
                }
            } else {
                console.log(`RFC no encontrado en Mongo: ${rfc}`);
            }

            const update = {
                $set: {
                    SINDICATO: {
                        AFILIADO: true,
                        DELEGACION: delegacion,
                        DELEGADO: delegado,
                        FECHA_AFILIACION: ''
                    }
                }
            };

            let result;
            if (collectionName === 'PLANTILLA') {
                result = await plantilla.updateOne(query, update, { upsert: false });
            } else {
                result = await licencias.updateOne(query, update);
            }

            if (result.matchedCount > 0) {
                console.log(`Actualizado en ${collectionName}: ${rfc} (Delegacion: ${delegacion}, Delegado: ${delegado})`);
            } else {
                rfcsNoEncontrados.push({ RFC: rfc, NOMBRE: nombre, NUMPLA: numpla });
                console.log(`RFC no encontrado: ${rfc}`);
            }
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