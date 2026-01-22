/* CHANGE THIS IN THE onOpen FUNCTION FOR IT TO WORK, OTHERWISE IT'LL NOT APPEAR ON THE MENU
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('✅ Prepare New Words for Export', 'prepareAnkiExport')
    .addSeparator() // Una línea separadora visual
    .addItem('🎧 Rellenar Audios Faltantes', 'generateMissingAudios') // <--- NUEVO
    .addToUi();
}
*/

// === 🛠️ HERRAMIENTA: RELLENAR AUDIOS FALTANTES ===

function generateMissingAudios() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Anki');
  
  if (!sheet) { ui.alert("No se encontró la hoja 'Anki'"); return; }

  // Indices de columnas (Basado en A=0, B=1...)
  // C=2 (Word), E=4 (Example), J=9 (Audio Word), L=11 (Audio Sentence)
  const WORD_IDX = 2;
  const EXAMPLE_IDX = 4;
  const AUDIO_WORD_IDX = 9;
  const AUDIO_SENT_IDX = 11;

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  let generatedCount = 0;
  
  // Notificación inicial
  ss.toast("Analizando filas en busca de audios faltantes...", "Iniciando", 5);

  // Empezamos en i=1 para saltar el encabezado
  for (let i = 1; i < data.length; i++) {
    const word = data[i][WORD_IDX];
    
    // Si no hay palabra, saltamos la fila
    if (!word) continue;

    const rowNumber = i + 1;
    let rowUpdated = false;

    // 1. VERIFICAR AUDIO DE PALABRA
    if (data[i][AUDIO_WORD_IDX] === "") {
      try {
        console.log(`🎤 Generando Audio Palabra para: "${word}"...`);
        const filename = `word_${cleanFilename(word)}.mp3`;
        // Llamamos a tu función existente
        const result = callOpenAITTS(word, filename);
        
        if (result) {
          sheet.getRange(rowNumber, AUDIO_WORD_IDX + 1).setValue(result);
          generatedCount++;
          rowUpdated = true;
        }
      } catch (e) {
        console.error(`Error palabra fila ${rowNumber}: ${e.message}`);
      }
    }

    // 2. VERIFICAR AUDIO DE FRASE
    if (data[i][AUDIO_SENT_IDX] === "") {
      const rawExample = data[i][EXAMPLE_IDX];
      
      // Solo generamos si hay un ejemplo escrito
      if (rawExample && rawExample.toString().trim() !== "") {
        try {
          // Limpiamos el formato Cloze {{c1::word}} -> word
          const cleanText = rawExample.toString().replace(/\{\{c\d+::(.*?)\}\}/g, '$1');
          
          console.log(`🗣️ Generando Audio Frase para fila ${rowNumber}...`);
          const filename = `sent_${cleanFilename(word)}.mp3`;
          
          // Llamamos a tu función existente
          const result = callOpenAITTS(cleanText, filename);
          
          if (result) {
            sheet.getRange(rowNumber, AUDIO_SENT_IDX + 1).setValue(result);
            generatedCount++;
            rowUpdated = true;
          }
        } catch (e) {
          console.error(`Error frase fila ${rowNumber}: ${e.message}`);
        }
      }
    }

    // Forzamos guardar cambios en el Sheet cada 5 actualizaciones para no perder datos si falla
    if (rowUpdated && generatedCount % 5 === 0) {
      SpreadsheetApp.flush();
      ss.toast(`Generados ${generatedCount} audios hasta ahora...`, "Progreso");
    }
  }

  ui.alert('✅ Proceso Finalizado', `Se generaron ${generatedCount} audios nuevos que faltaban.`, ui.ButtonSet.OK);
}
