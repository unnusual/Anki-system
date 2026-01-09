// === CONFIGURATION ===
const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA' // Tu carpeta de Drive
};

// === 1. INITIALIZATION ===
function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema listo: Gemini 2.5 + OpenAI TTS vinculados (V2.1 Pronunciation Support).');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

// === 2. MAIN PROCESSOR ===
function processFormSubmission(e) {
  try {
    const wordData = extractFormData(e);
    if (!wordData.palabra) return;

    const sheet = ensureAnkiSheet();
    // Verificamos duplicados
    const existingWords = sheet.getRange("C:C").getValues().flat().map(w => w.toString().toLowerCase());
    
    // NOTA: Permitimos duplicados si el modo es distinto (opcional), 
    // pero por seguridad mantenemos el bloqueo de duplicados estrictos por ahora.
    if (existingWords.includes(wordData.palabra.toLowerCase())) {
      console.log(`⏭️ Saltando duplicado: ${wordData.palabra}`);
      return;
    }

    // 1. Obtener datos de Gemini (Lógica Condicional aplicada dentro)
    const enriched = callGeminiAPI(wordData);
    
    // 2. Generar Audio para la palabra
    // Usamos el mismo TTS, pero si es pronunciación, la "palabra" base es la misma.
    const audioFilename = `${wordData.palabra.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_${Utilities.getUuid().substring(0, 4)}.mp3`;
    enriched.audio = callOpenAITTS(wordData.palabra, audioFilename);

    // 3. Guardar en la hoja
    // Definimos la etiqueta basada en el modo
    enriched.tag = wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab';

    addToAnkiSheet(enriched);
    
  } catch (error) {
    console.error('❌ Error en el proceso:', error.toString());
  }
}

// === 3. OPENAI TTS LOGIC ===
function callOpenAITTS(text, filename) {
  const url = "https://api.openai.com/v1/audio/speech";
  const payload = {
    model: "tts-1",
    input: text,
    voice: "nova", 
    response_format: "mp3"
  };

  const options = {
    method: "post",
    headers: { "Authorization": "Bearer " + CONFIG.OPENAI_API_KEY },
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    console.error(`🚫 Error OpenAI TTS: ${response.getContentText()}`);
    return "Error Audio";
  }

  const blob = response.getBlob().setName(filename);
  const folder = DriveApp.getFolderById(CONFIG.AUDIO_FOLDER_ID);
  folder.createFile(blob);
  
  return `[sound:${filename}]`;
}

// === 4. BATCH PROCESSOR (Para palabras viejas) ===
function generateAudioForExistingRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Anki');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let audioIdx = headers.indexOf('Audio');
  if (audioIdx === -1) {
    audioIdx = 8; // Fallback
    sheet.getRange(1, audioIdx + 1).setValue('Audio').setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
  }

  const wordIdx = 2; // Columna C

  for (let i = 1; i < data.length; i++) {
    const word = data[i][wordIdx];
    const existingAudio = data[i][audioIdx];
    
    if (word && (!existingAudio || existingAudio === "" || existingAudio === "Error Audio")) {
      console.log(`🎙️ Generando audio para: ${word}`);
      const filename = `audio_${Utilities.getUuid().substring(0, 8)}.mp3`;
      try {
        const audioTag = callOpenAITTS(word, filename);
        sheet.getRange(i + 1, audioIdx + 1).setValue(audioTag);
        Utilities.sleep(500); 
      } catch (e) {
        console.error(`Error en fila ${i+1}: ${e}`);
      }
    }
  }
  
  try {
    SpreadsheetApp.getUi().alert('✅ Proceso de audio completado.');
  } catch (e) {
    console.log('✅ Proceso de audio completado (UI no disponible).');
  }
}

// === 5. GEMINI API CALL (MODIFICADO PARA PRONUNCIACIÓN) ===
function callGeminiAPI(wordData) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  // Lógica de Prompt Condicional
  let promptText = "";
  
  if (wordData.modo === 'Solo Pronunciación') {
    // PROMPT PARA PRONUNCIACIÓN
    promptText = `
      You are an expert phonetician and linguist. 
      Analyze the word/phrase: "${wordData.palabra}". 
      Context provided: "${wordData.contexto}".
      
      Your goal is to help a student master the pronunciation.
      Return a JSON object with these exact keys:
      - definition: Provide the IPA (International Phonetic Alphabet) transcription. If it's a heteronym (like 'record'), use the context to choose the right stress.
      - example: Provide a specific tip about the stress pattern, silent letters, or linking sounds (e.g., "Stress on the 1st syllable because it is a noun").
      - type: The grammatical type (Noun, Verb, etc.).
    `;
  } else {
    // PROMPT ESTÁNDAR (Vocabulario General)
    promptText = `
      As a lexicographer, define "${wordData.palabra}". 
      Context provided: "${wordData.contexto}". 
      Provide a clear English definition, one original usage example, and the correct grammatical type.
      Return a JSON object with keys: definition, example, type.
    `;
  }

  const payload = {
    contents: [{ 
      parts: [{ text: promptText }] 
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json",
      responseSchema: {
        type: "object",
        properties: {
          definition: { type: "string" },
          example: { type: "string" },
          type: { type: "string" }
        },
        required: ["definition", "example", "type"]
      }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });

  const json = JSON.parse(response.getContentText());
  const result = JSON.parse(json.candidates[0].content.parts[0].text);

  return {
    palabra: wordData.palabra,
    definicion: result.definition, // Aquí irá el IPA si es modo pronunciación
    ejemplo: result.example,       // Aquí irá el Tip de Acento si es modo pronunciación
    tipo: result.type,
    contexto: wordData.contexto,
    modo: wordData.modo
  };
}

// === 6. DATA MANAGEMENT & EXPORT ===

function extractFormData(e) {
  const vals = e.namedValues;
  return {
    palabra: vals['Palabra o frase que quieres aprender'] ? vals['Palabra o frase que quieres aprender'][0].trim() : '',
    contexto: vals['Contexto u oración donde la viste (opcional)'] ? vals['Contexto u oración donde la viste (opcional)'][0].trim() : '',
    tipo: vals['Tipo de palabra (opcional)'] ? vals['Tipo de palabra (opcional)'][0].trim() : '',
    // Capturamos el nuevo campo. Si está vacío, asumimos General.
    modo: vals['Modo de Estudio'] ? vals['Modo de Estudio'][0].trim() : 'Vocabulario General'
  };
}

function ensureAnkiSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Anki') || ss.insertSheet('Anki');
  
  // Agregamos 'Tags' al final de los headers
  const headers = ['ID', 'Date', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Imported', 'Audio', 'Tags'];
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
    sheet.setFrozenRows(1);
  } else {
    // Verificación defensiva de la columna Audio
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (currentHeaders.indexOf('Audio') === -1) {
       sheet.getRange(1, 9).setValue('Audio').setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
    }
  }
  return sheet;
}

function addToAnkiSheet(data) {
  const sheet = ensureAnkiSheet();
  sheet.appendRow([
    Utilities.getUuid().substring(0, 8),
    new Date().toLocaleDateString(),
    data.palabra,
    data.definicion,
    data.ejemplo,
    data.contexto,
    data.tipo,
    'NO',
    data.audio,
    data.tag // <--- Nueva columna J
  ]);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('Prepare New Words for Export', 'prepareAnkiExport')
    .addItem('Generate Audio for Missing Rows', 'generateAudioForExistingRows')
    .addToUi();
}

function prepareAnkiExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Anki');
  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const statusIdx = headers.indexOf('Imported');
  
  const newWords = data.filter((row, index) => index > 0 && row[statusIdx] === 'NO');
  
  if (newWords.length === 0) {
    SpreadsheetApp.getUi().alert('No new words to export.');
    return;
  }

  let exportSheet = ss.getSheetByName('Anki_Export') || ss.insertSheet('Anki_Export');
  exportSheet.clear();
  
  const exportHeaders = ['ID', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Audio', 'Tags'];
  exportSheet.getRange(1, 1, 1, exportHeaders.length).setValues([exportHeaders]).setFontWeight('bold');
  
  // Ajustamos el mapeo. Asumiendo que 'Tags' está en la columna 10 (J), índice 9
  // ID(0), Word(2), Def(3), Ex(4), Context(5), Type(6), Audio(8), Tags(9)
  const rowsToExport = newWords.map(r => [r[0], r[2], r[3], r[4], r[5], r[6], r[8], r[9]]);
  exportSheet.getRange(2, 1, rowsToExport.length, exportHeaders.length).setValues(rowsToExport);

  for (let i = 2; i <= sourceSheet.getLastRow(); i++) {
    if (sourceSheet.getRange(i, statusIdx + 1).getValue() === 'NO') {
      sourceSheet.getRange(i, statusIdx + 1).setValue('YES');
    }
  }

  exportSheet.activate();
  SpreadsheetApp.getUi().alert(`✅ Export preparado con ${newWords.length} palabras.`);
}
