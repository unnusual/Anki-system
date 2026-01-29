/* CHANGE THIS IN THE onOpen FUNCTION FOR IT TO WORK, OTHERWISE IT'LL NOT APPEAR ON THE MENU
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('✅ Prepare New Words for Export', 'prepareAnkiExport')
    .addSeparator() // Una línea separadora visual
    .addItem('🎧 Rellenar Audios Faltantes', 'generateMissingAudios') // <--- NUEVO
    .addToUi();
}
*/

// === 🛠️ HERRAMIENTA: RELLENAR AUDIOS FALTANTES (VERSION CHIRP 3 HD) ===
function generateMissingAudios() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Anki');
  
  if (!sheet) { ui.alert("No se encontró la hoja 'Anki'"); return; }

  // Ajusta estos índices según tu realidad (A=0, B=1, etc.)
  const WORD_IDX = 2;       // Columna C
  const EXAMPLE_IDX = 4;    // Columna E
  const AUDIO_WORD_IDX = 9; // Columna J
  const AUDIO_SENT_IDX = 11;// Columna L

  const data = sheet.getDataRange().getValues();
  let generatedCount = 0;
  
  // LÍMITE DE SEGURIDAD: Procesar máximo 15 filas por ejecución para evitar Timeouts de 6 min
  const MAX_ROWS_PER_RUN = 15;
  let rowsProcessedInThisRun = 0;

  ss.toast("Analizando audios con Kore y Leda...", "Iniciando", 5);

  for (let i = 1; i < data.length && rowsProcessedInThisRun < MAX_ROWS_PER_RUN; i++) {
    const word = data[i][WORD_IDX];
    if (!word) continue;

    const rowNumber = i + 1;
    let rowUpdated = false;

    // 1. VERIFICAR AUDIO DE PALABRA (Usando Kore)
    if (data[i][AUDIO_WORD_IDX] === "") {
      try {
        console.log(`🎤 Kore generando palabra: "${word}"...`);
        const filename = `word_${cleanFilename(word)}.mp3`;
        
        // PASO IMPORTANTE: Enviamos "word" como tercer parámetro
        const result = callGoogleTTS(word, filename, "word");
        
        if (result) {
          sheet.getRange(rowNumber, AUDIO_WORD_IDX + 1).setValue(result);
          generatedCount++;
          rowUpdated = true;
        }
        Utilities.sleep(1500); // Pausa para no saturar la API
      } catch (e) {
        console.error(`Error palabra fila ${rowNumber}: ${e.message}`);
      }
    }

    // 2. VERIFICAR AUDIO DE FRASE (Usando Leda)
    if (data[i][AUDIO_SENT_IDX] === "") {
      const rawExample = data[i][EXAMPLE_IDX];
      
      if (rawExample && rawExample.toString().trim() !== "") {
        try {
          const cleanText = rawExample.toString().replace(/\{\{c\d+::(.*?)\}\}/g, '$1');
          
          console.log(`🗣️ Leda generando frase para: "${word}"...`);
          const filename = `sent_${cleanFilename(word)}.mp3`;
          
          // PASO IMPORTANTE: Enviamos "sentence" como tercer parámetro
          const result = callGoogleTTS(cleanText, filename, "sentence");
          
          if (result) {
            sheet.getRange(rowNumber, AUDIO_SENT_IDX + 1).setValue(result);
            generatedCount++;
            rowUpdated = true;
          }
          Utilities.sleep(1500); // Pausa para no saturar la API
        } catch (e) {
          console.error(`Error frase fila ${rowNumber}: ${e.message}`);
        }
      }
    }

    if (rowUpdated) {
      rowsProcessedInThisRun++;
      if (generatedCount % 3 === 0) {
        SpreadsheetApp.flush();
        ss.toast(`Generados ${generatedCount} audios...`, "Progreso");
      }
    }
  }

  const msg = (rowsProcessedInThisRun >= MAX_ROWS_PER_RUN) 
    ? `Límite de lote alcanzado (${MAX_ROWS_PER_RUN} filas). Se generaron ${generatedCount} audios. Ejecuta de nuevo para continuar.`
    : `Proceso finalizado. Se generaron ${generatedCount} audios nuevos.`;

  ui.alert('✅ Resultado del Lote', msg, ui.ButtonSet.OK);
}
