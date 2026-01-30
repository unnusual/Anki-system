// === 9. BATCH IMAGE FILLER (VERSIÓN FILTRADA) ===
function fillMissingImages() {
  console.log("🏁 INICIANDO PROCESO: Rellenar imágenes para Vocabulario General.");

  const sheet = ensureAnkiSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const wordCol = headers.indexOf('Word');
  const definitionCol = headers.indexOf('Definition');
  const exampleRawCol = headers.indexOf('Example_Sentence');
  const imageCol = headers.indexOf('Image');
  const tagsCol = headers.indexOf('Tags');

  if (wordCol === -1 || imageCol === -1 || tagsCol === -1) {
    SpreadsheetApp.getUi().alert("Error: No se encuentran las columnas Word, Image o Tags.");
    return;
  }

  const rowsToProcess = [];
  
  // Recorremos la hoja buscando vacíos que cumplan el criterio
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const word = row[wordCol];
    const currentImage = row[imageCol];
    const tag = row[tagsCol] ? row[tagsCol].toString().toLowerCase() : "";

    // CRITERIO: Sin imagen Y que el tag NO sea de pronunciación
    // Ajusta "general_vocab" si el nombre de tu tag es distinto
    if (word && !currentImage && tag.includes("general_vocab")) {
      rowsToProcess.push({
        rowIndex: i + 1,
        palabra: word,
        definicion: row[definitionCol],
        ejemplo_raw: row[exampleRawCol] || row[definitionCol]
      });
    }
  }

  if (rowsToProcess.length === 0) {
    console.log("✅ No hay palabras de vocabulario pendientes.");
    return;
  }

  console.log(`⏳ Procesando ${rowsToProcess.length} imágenes...`);

  // Límite de seguridad para evitar Timeouts de Google Scripts
  const LIMIT = 10; 
  const processedCount = Math.min(rowsToProcess.length, LIMIT);

  for (let i = 0; i < processedCount; i++) {
    const entry = rowsToProcess[i];
    try {
      const visualPrompt = generateVisualPrompt(entry);
      const imgFilename = `img_${cleanFilename(entry.palabra)}.png`;
      const generatedImageName = callOpenAIDalle(visualPrompt, imgFilename);

      if (generatedImageName) {
        sheet.getRange(entry.rowIndex, imageCol + 1).setValue(generatedImageName);
        console.log(`✅ [${i+1}/${processedCount}] Imagen creada: ${entry.palabra}`);
      }
      
      // Delay de cortesía para la API
      Utilities.sleep(3000); 
    } catch (err) {
      console.error(`❌ Error en "${entry.palabra}": ${err.message}`);
    }
  }
}
