// === CONFIGURATION ===
const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  
  // 🔴 IMPORTANTE: Reemplaza esto con el ID de texto de tu proyecto GCP (ej: 'anki-gen-ai-x82')
  GCP_PROJECT_ID: 'anki-gen-ai', 
  
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA', // Tu carpeta de Audio
  IMAGE_FOLDER_ID: '1TlEjDBQtyTYk0qBoalwkwCw3X8nA1Z4h'  // Tu carpeta de Imágenes (Link que enviaste)
};

// === 1. INITIALIZATION ===
function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V3.0 Enterprise Listo: Gemini + Vertex AI + OpenAI + Cloze Support.');
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

    // Usamos la nueva hoja V3 para soportar las nuevas columnas (Imágenes, Audio Contexto, Cloze)
    const sheet = ensureAnkiSheet(); 
    
    // Verificación de duplicados
    const existingWords = sheet.getRange("B:B").getValues().flat().map(w => w.toString().toLowerCase());
    if (existingWords.includes(wordData.palabra.toLowerCase())) {
      console.log(`⏭️ Saltando duplicado: ${wordData.palabra}`);
      return;
    }

    // A. CEREBRO: Gemini procesa todo (Definición, Cloze, Frecuencia, Prompt de Imagen)
    const enriched = callGeminiAnalyst(wordData);
    
    // B. VOZ: Generar Audios
    // 1. Audio de la palabra (Siempre)
    const wordFilename = `word_${cleanFilename(wordData.palabra)}.mp3`;
    enriched.audioWord = callOpenAITTS(wordData.palabra, wordFilename);

    // 2. Audio del Ejemplo (Solo si existe y NO es modo pronunciación)
    if (enriched.ejemplo_raw && wordData.modo !== 'Solo Pronunciación') {
       const sentenceFilename = `sent_${cleanFilename(wordData.palabra)}.mp3`;
       enriched.audioSentence = callOpenAITTS(enriched.ejemplo_raw, sentenceFilename);
    } else {
       enriched.audioSentence = "";
    }

    // C. VISIÓN: Generar Imagen con Vertex AI (Solo Vocabulario General)
    // Usamos tus créditos de GCP aquí.
    if (wordData.modo !== 'Solo Pronunciación' && enriched.image_prompt) {
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.png`;
      enriched.image = callVertexAIImage(enriched.image_prompt, imgFilename);
    } else {
      enriched.image = "";
    }

    // D. BASE DE DATOS: Guardar
    addToAnkiSheet(enriched);
    
  } catch (error) {
    console.error('❌ Error Critical:', error.toString());
  }
}

// === 3. GEMINI ANALYST (LOGIC CORE V3) ===
function callGeminiAnalyst(wordData) {
  // Usamos el modelo experimental o 1.5-flash que es rápido y soporta JSON complejo
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  let promptText = "";
  
  if (wordData.modo === 'Solo Pronunciación') {
    promptText = `
      Analyze the word: "${wordData.palabra}". Context: "${wordData.contexto}".
      Goal: Pronunciation mastery.
      Return JSON:
      - definition: IPA transcription.
      - example: A short technical tip about stress or sound linking.
      - example_raw: null
      - type: Grammatical type.
      - frequency_tag: "#Pronunciation"
      - image_prompt: null
    `;
  } else {
    // MODO VOCABULARIO (Con Cloze y Frecuencia)
    promptText = `
      Analyze the word: "${wordData.palabra}". Context: "${wordData.contexto}".
      Goal: Create an Anki card for an advanced learner.
      Return JSON:
      - definition: Clear English definition.
      - example: An original sentence using the word, but format it strictly as an Anki Cloze deletion (e.g., "The {{c1::apple}} is red").
      - example_raw: The same sentence but clean text (for Audio generation).
      - type: Grammatical type.
      - frequency_tag: The CEFR level (e.g., "#CEFR_B2", "#CEFR_C1") and a rarity tag if applicable (e.g., "#Academic", "#Slang").
      - image_prompt: A detailed prompt to generate a minimalist, vector-style illustration representing this word.
    `;
  }

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: { 
      responseMimeType: "application/json",
      responseSchema: {
        type: "object",
        properties: {
          definition: {type: "string"},
          example: {type: "string"},
          example_raw: {type: "string", nullable: true},
          type: {type: "string"},
          frequency_tag: {type: "string"},
          image_prompt: {type: "string", nullable: true}
        }
      }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload)
  });

  const result = JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);

  return {
    ...wordData, 
    definicion: result.definition,
    ejemplo: result.example,     // Versión Cloze {{c1::word}}
    ejemplo_raw: result.example_raw, // Versión Limpia para TTS
    tipo: result.type,
    tags: result.frequency_tag,
    image_prompt: result.image_prompt,
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 4. VERTEX AI (IMAGEN 3) - USANDO GCP CREDITS ===
function callVertexAIImage(prompt, filename) {
  if (!CONFIG.GCP_PROJECT_ID || CONFIG.GCP_PROJECT_ID.includes('TU_ID')) {
    console.error("GCP Project ID no configurado.");
    return "";
  }

  const location = 'us-central1'; 
  const modelId = 'imagegeneration@005'; // Modelo Imagen 3 (o el disponible)
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

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    
    if (response.getResponseCode() !== 200) {
      console.error("Vertex AI Error:", response.getContentText());
      return "";
    }

    const json = JSON.parse(response.getContentText());
    // Imagen 3 suele devolver bytesBase64Encoded
    if (json.predictions && json.predictions[0] && json.predictions[0].bytesBase64Encoded) {
      const base64Image = json.predictions[0].bytesBase64Encoded;
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Image), 'image/png', filename);
      
      const folder = DriveApp.getFolderById(CONFIG.IMAGE_FOLDER_ID);
      folder.createFile(blob);
      
      return `<img src="${filename}">`; 
    }
    return "";
  } catch (e) {
    console.error("Excepción Vertex AI:", e);
    return "";
  }
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
  if (response.getResponseCode() !== 200) return "Error Audio";

  const blob = response.getBlob().setName(filename);
  DriveApp.getFolderById(CONFIG.AUDIO_FOLDER_ID).createFile(blob);
  return `[sound:${filename}]`;
}

// === UTILS & SHEETS ===

function cleanFilename(text) {
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + Utilities.getUuid().substring(0,4);
}

function extractFormData(e) {
  const vals = e.namedValues;
  return {
    palabra: vals['Palabra o frase que quieres aprender'] ? vals['Palabra o frase que quieres aprender'][0].trim() : '',
    contexto: vals['Contexto u oración donde la viste (opcional)'] ? vals['Contexto u oración donde la viste (opcional)'][0].trim() : '',
    tipo: vals['Tipo de palabra (opcional)'] ? vals['Tipo de palabra (opcional)'][0].trim() : '',
    modo: vals['Modo de Estudio'] ? vals['Modo de Estudio'][0].trim() : 'Vocabulario General'
  };
}

// Crea la Hoja V3 preparada para el template "Beautiful Anki"
function ensureAnkiSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Anki_V3'); 
  
  if (!sheet) {
    sheet = ss.insertSheet('Anki_V3');
    // Headers alineados con los campos del Note Type
    // ID | Word | Definition | Example (Cloze) | Type | Tags | Image | Audio_Word | Audio_Sentence
    const headers = ['ID', 'Word', 'Definition', 'Example', 'Type', 'Tags', 'Image', 'Audio_Word', 'Audio_Sentence'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#2E7D32').setFontColor('white'); // Verde para diferenciar V3
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function addToAnkiSheet(data) {
  const sheet = ensureAnkiSheet();
  
  // Limpieza de tags: Combinamos el modo + frecuencia + duplicados
  let tagsClean = `${data.tag_mode} ${data.tags || ''}`;
  tagsClean = tagsClean.replace(/\s+/g, ' ').trim(); // Quitar espacios extra

  sheet.appendRow([
    Utilities.getUuid().substring(0, 8),
    data.palabra,
    data.definicion,
    data.ejemplo, // Contiene {{c1::word}}
    data.tipo,
    tagsClean,
    data.image,
    data.audioWord,
    data.audioSentence
  ]);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki V3 Enterprise')
    .addItem('Prepare V3 Export', 'prepareAnkiExportV3')
    .addToUi();
}

function prepareAnkiExportV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Anki_V3');
  if (!sourceSheet) return;

  const data = sourceSheet.getDataRange().getValues();
  // Filtramos los que no tengan ID en el export (en este caso exportamos todo lo nuevo manual)
  // O podemos crear lógica de "Exported?" si quieres. 
  // Para V3 simplificado: Exporta todo a una hoja limpia CSV.
  
  let exportSheet = ss.getSheetByName('Anki_Import_Ready') || ss.insertSheet('Anki_Import_Ready');
  exportSheet.clear();
  
  // Headers exactos para Anki
  // Nota: Quitamos el ID para el import, o lo dejamos si usas un addon de actualización.
  // Asumiremos importación estándar con ID como primer campo.
  exportSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  exportSheet.activate();
  SpreadsheetApp.getUi().alert(`✅ Datos listos en 'Anki_Import_Ready'. Exporta esta hoja como CSV.`);
}
