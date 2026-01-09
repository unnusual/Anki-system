// === CONFIGURATION ===
const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  GCP_PROJECT_ID: 'anki-gen-ai', 
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA', // Tu carpeta de Audio
  IMAGE_FOLDER_ID: '1TlEjDBQtyTYk0qBoalwkwCw3X8nA1Z4h'  // Tu carpeta de Imágenes
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

// === 2. MAIN PROCESSOR (SAFE MODE) ===
function processFormSubmission(e) {
  console.log("🏁 INICIANDO PROCESO BLINDADO...");
  
  // 1. Extracción de datos
  let wordData;
  try {
    wordData = extractFormData(e);
    if (!wordData.palabra) { console.warn("⚠️ No se detectó palabra."); return; }
    console.log(`📌 Procesando: "${wordData.palabra}"`);
  } catch (err) {
    console.error("❌ Error extrayendo datos:", err);
    return;
  }

  // 2. CEREBRO (Gemini) - Crítico (Si falla esto, paramos)
  let enriched;
  try {
    enriched = callGeminiAnalyst(wordData);
    console.log("✅ Gemini: Datos generados correctamente.");
  } catch (err) {
    console.error("❌ ERROR CRÍTICO GEMINI:", err);
    return; // Sin cerebro no hay tarjeta
  }

  // 3. VOZ (OpenAI) - Opcional
  try {
    const wordFilename = `word_${cleanFilename(wordData.palabra)}.mp3`;
    enriched.audioWord = callOpenAITTS(wordData.palabra, wordFilename);
    
    if (enriched.ejemplo_raw && wordData.modo !== 'Solo Pronunciación') {
       const sentenceFilename = `sent_${cleanFilename(wordData.palabra)}.mp3`;
       enriched.audioSentence = callOpenAITTS(enriched.ejemplo_raw, sentenceFilename);
    } else {
       enriched.audioSentence = "";
    }
    console.log("✅ Audio: Generado.");
  } catch (err) {
    console.error("⚠️ Error en Audio (Continuando sin audio):", err);
    enriched.audioWord = "";
    enriched.audioSentence = "";
  }

  // 4. VISIÓN (Vertex AI) - Opcional y Riesgoso
  // Aquí es donde sospecho que ocurría el crash anterior
  try {
    if (wordData.modo !== 'Solo Pronunciación' && enriched.image_prompt) {
      console.log("🎨 Iniciando generación de imagen (Vertex AI)...");
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.png`;
      
      // Verificamos acceso a carpeta antes de llamar a la IA
      const folder = DriveApp.getFolderById(CONFIG.IMAGE_FOLDER_ID); 
      
      enriched.image = callVertexAIImage(enriched.image_prompt, imgFilename);
      console.log("✅ Imagen: Generada y guardada.");
    } else {
      enriched.image = "";
    }
  } catch (err) {
    console.error("⚠️ Error en Imagen (Continuando sin imagen):", err.toString());
    enriched.image = ""; // Dejamos la imagen vacía para no romper la tarjeta
  }

  // 5. BASE DE DATOS (Sheets)
  try {
    addToAnkiSheet(enriched);
    console.log("🎉 ÉXITO: Tarjeta guardada en Sheets.");
  } catch (err) {
    console.error("❌ Error guardando en Sheets:", err);
  }
}

// === 3. GEMINI ANALYST (LOGIC CORE V3) - VERSIÓN 2026 STABLE ===
function callGeminiAnalyst(wordData) {
  // ✅ USAMOS EL MODELO ESTABLE QUE APARECIÓ EN TU LISTA
  // "gemini-2.5-flash" es rápido, inteligente y estable.
  const modelVersion = 'gemini-2.5-flash'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  let promptText = "";
  
  // Prompt ajustado: Directo y sin espacio para alucinaciones
  if (wordData.modo === 'Solo Pronunciación') {
    promptText = `
      You are a linguistic database. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Task: Provide pronunciation data.
      Output strictly valid JSON.
      JSON Schema:
      {
        "definition": "IPA transcription only",
        "example": "1 sentence tip about pronunciation/stress",
        "example_raw": null,
        "type": "Part of speech",
        "frequency_tag": "#Pronunciation",
        "image_prompt": null
      }
    `;
  } else {
    promptText = `
      You are a linguistic database. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Task: Create data for an Anki card (Advanced English).
      Output strictly valid JSON.
      JSON Schema:
      {
        "definition": "Concise definition (max 15 words).",
        "example": "Sentence with Anki cloze format: 'The {{c1::word}} ...'",
        "example_raw": "The same sentence as plain text.",
        "type": "Part of speech (e.g. noun, verb).",
        "frequency_tag": "CEFR Level (e.g. #CEFR_C1).",
        "image_prompt": "Minimalist vector illustration description."
      }
    `;
  }

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: { 
      responseMimeType: "application/json",
      temperature: 0.1,    // 🥶 Temperatura muy baja = Máxima precisión, cero creatividad loca
    }
  };

  const options = {
    method: 'post', 
    contentType: 'application/json', 
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);

  // Validación de respuesta HTTP
  if (response.getResponseCode() !== 200) {
    console.error(`❌ Error Gemini API (${response.getResponseCode()}): ${response.getContentText()}`);
    throw new Error(`Gemini Falló: ${response.getContentText()}`);
  }

  try {
    const jsonResponse = JSON.parse(response.getContentText());
    
    // Validación de estructura de respuesta de Gemini
    if (!jsonResponse.candidates || !jsonResponse.candidates[0].content) {
      console.warn("⚠️ Respuesta vacía de Gemini:", response.getContentText());
      throw new Error("Gemini no devolvió contenido.");
    }

    const textContent = jsonResponse.candidates[0].content.parts[0].text;
    const result = JSON.parse(textContent);

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
  } catch (e) {
    console.error("❌ Error procesando JSON final:", e);
    console.error("Respuesta cruda recibida:", response.getContentText());
    throw e;
  }
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

// === DEBUGGING / TEST ===
function testManualSubmission() {
  // Simulamos los datos que enviaría el formulario
  const mockEvent = {
    namedValues: {
      'Palabra o frase que quieres aprender': ['ephemeral'], // Cambia esto para probar otras palabras
      'Contexto u oración donde la viste (opcional)': ['The ephemeral joys of childhood.'],
      'Tipo de palabra (opcional)': ['adjective'],
      'Modo de Estudio': ['Vocabulario General'] 
    }
  };

  console.log("🧪 Iniciando prueba manual...");
  
  // Llamamos a la función principal pasándole los datos falsos
  processFormSubmission(mockEvent);
}
