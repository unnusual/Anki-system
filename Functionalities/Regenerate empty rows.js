/*
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('Prepare New Words for Export', 'prepareAnkiExport')
    .addSeparator()
    .addItem('🚑 Regenerar Filas Vacías', 'regenerateEmptyRows')
    .addToUi();
}
*/

// === 🚑 CIRUGÍA: REGENERAR FILAS VACÍAS ===

function regenerateEmptyRows() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Anki');
  
  if (!sheet) { ui.alert("No se encontró la hoja 'Anki'"); return; }

  // CONFIGURACIÓN DE RESCATE
  // Cambia esto a 'Solo Pronunciación' si prefieres que el arreglo sea en ese modo.
  const RESCUE_MODE = 'Solo Pronunciación'; 
  
  // Índices (A=0, C=2 Word, D=3 Definition, F=5 Context, I=8 Tags)
  const COL_WORD = 2; 
  const COL_DEF = 3;
  const COL_CONTEXT = 5;

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  let fixedCount = 0;

  ss.toast("Escanando filas corruptas...", "🚑 Iniciando", 5);

  // Recorremos fila por fila (saltando encabezado)
  for (let i = 1; i < data.length; i++) {
    const word = data[i][COL_WORD];
    const definition = data[i][COL_DEF];
    const context = data[i][COL_CONTEXT];

    // CRITERIO: Hay palabra, pero NO hay definición (celda vacía)
    if (word && (!definition || definition.toString().trim() === "")) {
      
      const rowNum = i + 1;
      console.log(`🔧 Reparando fila ${rowNum}: "${word}"...`);
      ss.toast(`Reparando: ${word}`, "Trabajando");

      try {
        // 1. Preparamos los datos simulados
        const mockData = {
          palabra: word,
          contexto: context || `Learning the word ${word}`, // Fallback si no hay contexto
          modo: RESCUE_MODE
        };

        // 2. Llamamos a Gemini (Tu función corregida V4.4 ya maneja esto)
        const enriched = callGeminiAnalyst(mockData);

        // 3. Generamos Audios si no existen
        // (Reutilizamos la lógica de 'generateMissingAudios' implícitamente aquí o simplemente los generamos de nuevo para asegurar)
        const audioWordFilename = `word_${cleanFilename(word)}.mp3`;
        const audioWordPath = callOpenAITTS(word, audioWordFilename);
        
        let audioSentPath = "";
        if (enriched.ejemplo_raw) {
           const audioSentFilename = `sent_${cleanFilename(word)}.mp3`;
           audioSentPath = callOpenAITTS(enriched.ejemplo_raw, audioSentFilename);
        }

        // 4. Escribimos directamente en las celdas (¡Cirugía precisa!)
        // Columnas: D(Def), E(Ex), G(Type), I(Tags), J(AudioW), L(AudioS)
        // Nota: Ajusta los índices según tu hoja real. 
        // D=4 (col 4), E=5, G=7, I=9, J=10, L=12 (en notación getRange que empieza en 1)
        
        sheet.getRange(rowNum, 4).setValue(enriched.definicion); // D
        sheet.getRange(rowNum, 5).setValue(enriched.ejemplo);    // E
        sheet.getRange(rowNum, 7).setValue(enriched.tipo);       // G
        sheet.getRange(rowNum, 9).setValue(enriched.tag_mode);   // I
        sheet.getRange(rowNum, 10).setValue(audioWordPath);      // J
        sheet.getRange(rowNum, 12).setValue(audioSentPath);      // L

        fixedCount++;
        SpreadsheetApp.flush(); // Guardar cambios inmediatos

      } catch (e) {
        console.error(`❌ Falló reparación de ${word}: ${e.message}`);
      }
    }
  }

  const msg = fixedCount > 0 
    ? `✅ Se repararon ${fixedCount} filas vacías exitosamente.`
    : `👍 No se encontraron filas vacías para reparar.`;
  
  ui.alert(msg);
}
