// script.js

document.getElementById('combineButton').addEventListener('click', combineFiles);

function combineFiles() {
    const excelFile = document.getElementById('excelFile').files[0];
    const wordFile = document.getElementById('wordFile').files[0];
    const message = document.getElementById('message');

    if (!excelFile || !wordFile) {
        message.textContent = 'Por favor, cargue ambos archivos.';
        return;
    }

    if (excelFile.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        wordFile.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        message.textContent = 'Asegúrese de cargar un archivo .xlsx y un archivo .docx.';
        return;
    }

    message.textContent = 'Procesando archivos...';

    // Leer archivo Excel
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        // Leer y procesar archivo Word
        const readerWord = new FileReader();
        readerWord.onload = (e) => {
            const doc = new Docxtemplater().loadZip(new PizZip(e.target.result));
            
            // Llenar plantillas en el archivo Word con datos de Excel
            excelData.forEach((row, index) => {
                doc.setData(row);

                try {
                    doc.render();
                } catch (error) {
                    console.error("Error en renderización de plantilla", error);
                    message.textContent = "Error al procesar la plantilla del archivo Word.";
                    return;
                }

                // Crear archivo de salida
                const blob = doc.getZip().generate({
                    type: "blob",
                    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                });

                const outputFilename = `documento_generado_${index + 1}.docx`;
                saveAs(blob, outputFilename);
            });

            message.textContent = "Documentos generados correctamente.";
        };
        readerWord.readAsArrayBuffer(wordFile);
    };
    reader.readAsArrayBuffer(excelFile);
}
