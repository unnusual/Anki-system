// === CONFIGURATION ===
const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  GCP_PROJECT_ID: 'anki-gen-ai', 
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA', 
  IMAGE_FOLDER_ID: '1TlEjDBQtyTYk0qBoalwkwCw3X8nA1Z4h'  
};

// === 1. INITIALIZATION ===
function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V3.1 Classic Listo: Estética original con Media al final.');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(ss).onFormSubmit().create();
}

// === HELPER: GUARDADO ROBUSTO (API AVANZADA) ===
function saveFileToDrive(blob, filename, folderId) {
  try {
    const fileMetadata = {
      title: filename,
      parents: [{id: folderId}],
      mimeType: blob.getContentType()
    };
    Drive.Files.insert(fileMetadata, blob);
    return true; 
  } catch (e) {
    console.error(`❌ Error guardando archivo ${filename}: ${e.message}`);
    throw e;
  }
}

// === 2. MAIN PROCESSOR ===
function processFormSubmission(e) {
  console.log("🏁 INICIANDO PROCESO V3.1 CLASSIC...");
  
  // 1. Extracción
  let wordData;
  try {
    wordData = extractFormData(e);
    if (!wordData.palabra) { console.warn("⚠️ No se detectó palabra."); return; }
    console.log(`📌 Procesando: "${wordData.palabra}"`);
  } catch (err) { console.error("❌ Error Data:", err); return; }

  // 1.5 VALIDACIÓN DE DUPLICADOS (Ajustado a formato Clásico)
  const sheet = ensureAnkiSheet();
  // En formato Clásico (ID, Date, Word...), la palabra está en la Columna C (3)
  const existingWords = sheet.getRange("C:C").getValues().flat()
    .filter(cell => cell !== "") 
    .map(w => w.toString().toLowerCase());

  if (existingWords.includes(wordData.palabra.toLowerCase())) {
    console.warn(`⏭️ DUPLICADO: "${wordData.palabra}".`);
    return;
  }

  // 2. CEREBRO (Gemini 2.5)
  let enriched;
  try {
    enriched = callGeminiAnalyst(wordData);
    console.log("✅ Gemini: Datos listos.");
  } catch (err) { console.error("❌ ERROR GEMINI:", err); return; }

  // 3. AUDIO (OpenAI - Dual Channel)
  try {
    // 3.1 Audio Palabra
    const wordFilename = `word_${cleanFilename(wordData.palabra)}.mp3`;
    enriched.audioWord = callOpenAITTS(wordData.palabra, wordFilename);
    
    // 3.2 Audio Frase (Si existe y no es solo pronunciación)
    if (enriched.ejemplo_raw && wordData.modo !== 'Solo Pronunciación') {
       console.log("🔹 Generando audio frase...");
       const sentenceFilename = `sent_${cleanFilename(wordData.palabra)}.mp3`;
       enriched.audioSentence = callOpenAITTS(enriched.ejemplo_raw, sentenceFilename);
    } else {
       enriched.audioSentence = "";
    }
  } catch (err) {
    console.error("⚠️ Error Audio:", err);
    enriched.audioWord = ""; enriched.audioSentence = "";
  }

  // 4. IMAGEN (Vertex AI)
  try {
    if (wordData.modo !== 'Solo Pronunciación' && enriched.image_prompt) {
      console.log("🎨 Generando imagen...");
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.png`;
      enriched.image = callVertexAIImage(enriched.image_prompt, imgFilename);
    } else {
      enriched.image = "";
    }
  } catch (err) {
    console.error("⚠️ Error Imagen:", err.toString());
    enriched.image = ""; 
  }

  // 5. GUARDAR
  try {
    addToAnkiSheet(enriched);
    console.log("🎉 ÉXITO TOTAL: Tarjeta creada en hoja 'Anki'.");
  } catch (err) { console.error("❌ Error Sheets:", err); }
}

// === 3. GEMINI ANALYST (2.5 FLASH) ===
function callGeminiAnalyst(wordData) {
  const modelVersion = 'gemini-2.5-flash'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  let promptText = "";
  if (wordData.modo === 'Solo Pronunciación') {
    promptText = `
      You are a linguistic engine. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Task: Pronunciation data. Output JSON.
      {
        "definition": "IPA transcription",
        "example": "Tip about stress/linking",
        "example_raw": null,
        "type": "Part of speech",
        "frequency_tag": "#Pronunciation",
        "image_prompt": null
      }
    `;
  } else {
    promptText = `
      You are a linguistic engine. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Task: Create Anki card. Output JSON.
      {
        "definition": "Concise definition (max 15 words).",
        "example": "Sentence with Anki cloze: 'The {{c1::word}} ...'",
        "example_raw": "Same sentence plain text for Audio TTS.",
        "type": "Part of speech.",
        "frequency_tag": "CEFR Level (e.g. #CEFR_B2).",
        "image_prompt": "Minimalist vector illustration description, flat design, white background."
      }
    `;
  }

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
  };

  const options = {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) throw new Error(response.getContentText());

  const result = JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);

  return {
    ...wordData, 
    definicion: result.definition,
    ejemplo: result.example,     
    ejemplo_raw: result.example_raw, 
    tipo: result.type,
    tags: result.frequency_tag,
    image_prompt: result.image_prompt,
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 4. VERTEX AI (IMAGEN) ===
function callVertexAIImage(prompt, filename) {
  if (!CONFIG.GCP_PROJECT_ID) return "";
  const location = 'us-central1'; 
  const modelId = 'imagegeneration@005'; 
  const endpoint = `https://${location}-aiplatform.googleapis.com/v1/projects/${CONFIG.GCP_PROJECT_ID}/locations/${location}/publishers/google/models/${modelId}:predict`;

  const payload = {
    instances: [{ prompt: prompt + ", minimalist vector art, flat design, white background, high quality" }],
    parameters: { sampleCount: 1, aspectRatio: "1:1" }
  };

  const options = {
    method: 'post',
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }, 
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  if (response.getResponseCode() !== 200) { console.error("Vertex AI Error:", response.getContentText()); return ""; }

  const json = JSON.parse(response.getContentText());
  if (json.predictions && json.predictions[0] && json.predictions[0].bytesBase64Encoded) {
    const blob = Utilities.newBlob(Utilities.base64Decode(json.predictions[0].bytesBase64Encoded), 'image/png', filename);
    saveFileToDrive(blob, filename, CONFIG.IMAGE_FOLDER_ID);
    return `<img src="${filename}">`; 
  }
  return "";
}

// === 5. OPENAI TTS ===
function callOpenAITTS(text, filename) {
  if (!text) return "";
  const url = "https://api.openai.com/v1/audio/speech";
  const payload = { model: "tts-1", input: text, voice: "nova", response_format: "mp3" };
  const options = {
    method: "post", headers: { "Authorization": "Bearer " + CONFIG.OPENAI_API_KEY },
    contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) return "";

  const blob = response.getBlob().setName(filename);
  saveFileToDrive(blob, filename, CONFIG.AUDIO_FOLDER_ID);
  return `[sound:${filename}]`;
}

// === UTILS & SHEETS (ESTRUCTURA CLÁSICA) ===

function cleanFilename(text) {
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + Utilities.getUuid().substring(0,4);
}

function extractFormData(e) {
  if (!e || !e.namedValues) return { palabra: "TEST_CLASSIC", contexto: "Test context", modo: "Vocabulario General" };
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
  // Volvemos a la hoja original 'Anki'
  let sheet = ss.getSheetByName('Anki'); 
  
  if (!sheet) {
    sheet = ss.insertSheet('Anki');
  }

  // Definimos los Headers Clásicos + Media al Final
  // A(1)  B(2)  C(3)  D(4)        E(5)     F(6)     G(7)  H(8)      I(9)  J(10)       K(11)  L(12)
  const headers = ['ID', 'Date', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Imported', 'Tags', 'Audio_Word', 'Image', 'Audio_Sentence'];
  
  // Si la primera fila está vacía o no coincide, la reescribimos para asegurar el orden
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === "" || firstCell !== 'ID') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
    sheet.setFrozenRows(1);
  } else {
    // Si ya existe, verificamos si tiene las nuevas columnas de media al final
    // Si no las tiene, las agregamos (Solo los headers)
    const lastCol = sheet.getLastColumn();
    if (lastCol < 12) {
       sheet.getRange(1, 10).setValue('Audio_Word').setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
       sheet.getRange(1, 11).setValue('Image').setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
       sheet.getRange(1, 12).setValue('Audio_Sentence').setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
    }
  }
  return sheet;
}

function addToAnkiSheet(data) {
  const sheet = ensureAnkiSheet();
  
  let tagsClean = `${data.tag_mode} ${data.tags || ''}`.replace(/\s+/g, ' ').trim();

  // Mapeo Clásico
  sheet.appendRow([
    Utilities.getUuid().substring(0, 8), // ID
    new Date().toLocaleDateString(),     // Date
    data.palabra,                        // Word
    data.definicion,                     // Definition
    data.ejemplo,                        // Example (Cloze)
    data.contexto,                       // Context
    data.tipo,                           // Type
    'NO',                                // Imported
    tagsClean,                           // Tags
    // --- MEDIA AL FINAL ---
    data.audioWord,                      // Audio_Word
    data.image,                          // Image
    data.audioSentence                   // Audio_Sentence
  ]);
}
