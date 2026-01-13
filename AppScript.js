// === CONFIGURATION ===
const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA' // Asegúrate que este ID sea correcto
};

// === 1. INITIALIZATION ===
function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V2.1 BLINDADO Listo: Drive API Advanced + Gemini 2.5');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

// === HELPER: GUARDADO ROBUSTO (LA SOLUCIÓN AL ERROR) ===
function saveFileToDrive(blob, filename, folderId) {
  try {
    const fileMetadata = {
      title: filename,
      parents: [{id: folderId}],
      mimeType: blob.getContentType()
    };
    // Usamos la API Avanzada en lugar de DriveApp
    const file = Drive.Files.insert(fileMetadata, blob);
    return file.alternateLink; 
  } catch (e) {
    console.error(`❌ Error guardando archivo ${filename}: ${e.message}`);
    throw e;
  }
}

// === 2. MAIN PROCESSOR ===
function processFormSubmission(e) {
  try {
    console.log("🏁 Iniciando proceso...");
    const wordData = extractFormData(e);
    if (!wordData.palabra) {
       console.warn("⚠️ No se detectó palabra.");
       return;
    }

    const sheet = ensureAnkiSheet();
    // Verificamos duplicados
    const existingWords = sheet.getRange("C:C").getValues().flat().map(w => w.toString().toLowerCase());
    
    if (existingWords.includes(wordData.palabra.toLowerCase())) {
      console.log(`⏭️ Saltando duplicado: ${wordData.palabra}`);
      return;
    }

    // 1. Obtener datos de Gemini
    console.log(`🔹 Consultando Gemini para: ${wordData.palabra}...`);
    const enriched = callGeminiAPI(wordData);
    
    // 2. Generar Audio (Usando el método blindado)
    console.log(`🔹 Generando audio...`);
    const audioFilename = `${cleanFilename(wordData.palabra)}.mp3`;
    enriched.audio = callOpenAITTS(wordData.palabra, audioFilename);

    // 3. Guardar en la hoja
    enriched.tag = wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab';

    addToAnkiSheet(enriched);
    console.log("🎉 ÉXITO: Palabra procesada y guardada.");
    
  } catch (error) {
    console.error('❌ Error Critical:', error.toString());
  }
}

// === 3. OPENAI TTS LOGIC (CORREGIDO) ===
function callOpenAITTS(text, filename) {
  if (!text) return "";
  
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
    return ""; // Retorna vacío si falla API, no rompe script
  }

  const blob = response.getBlob().setName(filename);
  
  // 👇 AQUÍ ESTÁ LA MAGIA: Usamos el helper avanzado
  try {
    saveFileToDrive(blob, filename, CONFIG.AUDIO_FOLDER_ID);
    return `[sound:${filename}]`;
  } catch (e) {
    console.error("Falló el guardado en Drive. Verifica el ID de carpeta.");
    return "";
  }
}

// === 5. GEMINI API CALL ===
function callGeminiAPI(wordData) {
  // Usamos el modelo estable 2.5
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  let promptText = "";
  if (wordData.modo === 'Solo Pronunciación') {
    promptText = `
      You are an expert phonetician. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Goal: Pronunciation mastery.
      Output JSON only:
      {
        "definition": "IPA transcription",
        "example": "Tip about stress/linking",
        "type": "Part of speech"
      }
    `;
  } else {
    promptText = `
      Define "${wordData.palabra}". Context: "${wordData.contexto}".
      Goal: Anki card.
      Output JSON only:
      {
        "definition": "Clear definition",
        "example": "Sentence example",
        "type": "Part of speech"
      }
    `;
  }

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error(`Gemini Error: ${response.getContentText()}`);
  }

  const json = JSON.parse(response.getContentText());
  const result = JSON.parse(json.candidates[0].content.parts[0].text);

  return {
    palabra: wordData.palabra,
    definicion: result.definition,
    ejemplo: result.example,
    tipo: result.type,
    contexto: wordData.contexto,
    modo: wordData.modo
  };
}

// === UTILS ===
function cleanFilename(text) {
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + Utilities.getUuid().substring(0,4);
}

function extractFormData(e) {
  if (!e || !e.namedValues) return { palabra: "TEST_WORD", contexto: "Test context", modo: "Vocabulario General" }; // Modo test
  const vals = e.namedValues;
  return {
    palabra: vals['Palabra o frase que quieres aprender'] ? vals['Palabra o frase que quieres aprender'][0].trim() : '',
    contexto: vals['Contexto u oración donde la viste (opcional)'] ? vals['Contexto u oración donde la viste (opcional)'][0].trim() : '',
    tipo: vals['Tipo de palabra (opcional)'] ? vals['Tipo de palabra (opcional)'][0].trim() : '',
    modo: vals['Modo de Estudio'] ? vals['Modo de Estudio'][0].trim() : 'Vocabulario General'
  };
}

function ensureAnkiSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Anki') || ss.insertSheet('Anki');
  const headers = ['ID', 'Date', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Imported', 'Audio', 'Tags'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
    sheet.setFrozenRows(1);
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
    data.tag 
  ]);
}
