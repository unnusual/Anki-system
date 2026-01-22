// === 🧹 HERRAMIENTAS DE LIMPIEZA (DRIVE CLEANER) ===

/**
 * MENU: Agrega esto dentro de la función onOpen() existente para ver los botones
 */
/* function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('✅ Prepare New Words for Export', 'prepareAnkiExport')
    .addSeparator()
    .addItem('🧹 SIMULAR Limpieza de Audios', 'simulateAudioCleanup') // Solo lista en consola
    .addItem('🗑️ EJECUTAR Limpieza de Audios', 'performAudioCleanup') // Borra archivos
    .addToUi();
}
*/

// 1. MODO SIMULACRO (Seguro: No borra nada, solo avisa)
function simulateAudioCleanup() {
  cleanDriveFolder(CONFIG.AUDIO_FOLDER_ID, true);
}

// 2. MODO EJECUCIÓN (Cuidado: Envía a la papelera)
function performAudioCleanup() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '⚠️ ¿Estás seguro?',
    'Esto enviará a la papelera todos los archivos de la carpeta de AUDIOS que no estén registrados en la hoja de cálculo actual.\n\nEsta acción no se puede deshacer automáticamente.',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    cleanDriveFolder(CONFIG.AUDIO_FOLDER_ID, false);
  }
}

// === LÓGICA PRINCIPAL DE LIMPIEZA ===
function cleanDriveFolder(folderId, isSimulation) {
  console.log(isSimulation ? "🕵️ INICIANDO SIMULACRO DE LIMPIEZA..." : "🗑️ INICIANDO LIMPIEZA REAL...");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Anki');
  
  if (!sheet) { console.error("No se encontró la hoja Anki"); return; }

  // 1. OBTENER LISTA DE ARCHIVOS VÁLIDOS (Los que están en el Excel)
  // Columnas J (Audio Word, índice 9) y L (Audio Sentence, índice 11)
  const data = sheet.getDataRange().getValues();
  const validFiles = new Set();
  
  // Empezamos en i=1 para saltar encabezados
  for (let i = 1; i < data.length; i++) {
    const wordAudio = data[i][9]; // Columna J
    const sentAudio = data[i][11]; // Columna L
    
    if (wordAudio) validFiles.add(wordAudio.toString().trim());
    if (sentAudio) validFiles.add(sentAudio.toString().trim());
  }
  
  console.log(`✅ Archivos válidos en la hoja: ${validFiles.size}`);

  // 2. ESCANEAR DRIVE
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  
  let deletedCount = 0;
  let keptCount = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Si el nombre del archivo NO está en la lista de válidos
    if (!validFiles.has(fileName)) {
      if (isSimulation) {
        console.log(`   [SIMULACRO] Se borraría: ${fileName}`);
      } else {
        try {
          file.setTrashed(true); // Lo envía a la papelera (se puede recuperar en 30 días)
          console.log(`   🗑️ ELIMINADO: ${fileName}`);
        } catch (e) {
          console.error(`   ❌ Error borrando ${fileName}: ${e.message}`);
        }
      }
      deletedCount++;
    } else {
      keptCount++;
    }
  }
  
  // 3. REPORTE FINAL
  const msg = isSimulation 
    ? `SIMULACRO: Se borrarían ${deletedCount} archivos basura. Se mantendrían ${keptCount} archivos correctos.`
    : `LIMPIEZA: Se eliminaron ${deletedCount} archivos basura. Quedan ${keptCount} archivos correctos.`;
    
  console.log("🏁 " + msg);
  SpreadsheetApp.getUi().alert(msg);
}
