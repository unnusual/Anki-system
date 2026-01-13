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
  console.log('🚀 Sistema V3.6 Clean DB Listo: Nombres limpios en Sheet, Etiquetas en Export.');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(ss).onFormSubmit().create();
}

// === HELPER: GUARDADO ROBUSTO ===
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
  console.log("🏁 INICIANDO PROCESO V3.6...");
  
  // 1. Extracción
  let wordData;
  try {
    wordData = extractFormData(e);
    if (!wordData.palabra) { console.warn("⚠️ No se detectó palabra."); return; }
    console.log(`📌 Procesando: "${wordData.palabra}"`);
  } catch (err) { console.error("❌ Error Data:", err); return; }

  // 1.5 Validación de Duplicados
  const sheet = ensureAnkiSheet();
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

  // 3. AUDIO (OpenAI)
  try {
    const wordFilename = `word_${cleanFilename(wordData.palabra)}.mp3`;
    enriched.audioWord = callOpenAITTS(wordData.palabra, wordFilename);
    
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
    console.log("🎉 ÉXITO TOTAL: Tarjeta guardada.");
  } catch (err) { console.error("❌ Error Sheets:", err); }
}

// === 3. GEMINI ANALYST (Mantiene lógica de seguridad V3.2) ===
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
        "frequency_tag": null,
        "image_prompt": null
      }
    `;
  } else {
    // Prompt optimizado para iconos vectoriales y seguridad
    promptText = `
      You are a linguistic engine. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Task: Create Anki card. Output JSON.
      
      CRITICAL FOR IMAGE_PROMPT: 
      - Create a SAFE, MINIMALIST vector icon description.
      - Abstract metaphors are preferred (e.g., "Shield" for Resilience).
      - PEOPLE: You MAY describe generic people/professions ONLY if describing a "vector icon" or "silhouette".
      - AVOID: "Photo", "Realistic", "Face", "Specific celebrity".
      - FORBIDDEN: Violence, weapons, blood, storms, disasters.

      JSON Schema:
      {
        "definition": "Concise definition (max 15 words).",
        "example": "Sentence with Anki cloze: 'The {{c1::word}} ...'",
        "example_raw": "Same sentence plain text for Audio TTS.",
        "type": "Part of speech.",
        "frequency_tag": "Thematic tag (e.g. #Business) or null. No CEFR.",
        "image_prompt": "Safe, minimalist vector icon description."
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
    tags: result.frequency_tag || null,
    image_prompt: result.image_prompt,
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 4. VERTEX AI (MODIFICADO: Retorna solo filename) ===
function callVertexAIImage(prompt, filename) {
  if (!CONFIG.GCP_PROJECT_ID) return "";
  const location = 'us-central1'; 
  const modelId = 'imagegeneration@005'; 
  const endpoint = `https://${location}-aiplatform.googleapis.com/v1/projects/${CONFIG.GCP_PROJECT_ID}/locations/${location}/publishers/google/models/${modelId}:predict`;

  const systemPrompt = ", vector art style, minimalist, white background.";
  const payload = {
    instances: [{ prompt: prompt + systemPrompt }],
    parameters: { sampleCount: 1, aspectRatio: "1:1" }
  };

  const options = {
    method: 'post', headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }, 
    contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    if (response.getResponseCode() === 400) {
       console.warn(`⚠️ Imagen Bloqueada. Continuando.`); return ""; 
    }
    if (response.getResponseCode() !== 200) return "";

    const json = JSON.parse(response.getContentText());
    if (json.predictions && json.predictions[0] && json.predictions[0].bytesBase64Encoded) {
      const blob = Utilities.newBlob(Utilities.base64Decode(json.predictions[0].bytesBase64Encoded), 'image/png', filename);
      saveFileToDrive(blob, filename, CONFIG.IMAGE_FOLDER_ID);
      
      // 👇 CAMBIO V3.6: Retornamos solo el nombre limpio del archivo
      return filename; 
    }
  } catch (e) { console.error("Excepción imagen:", e.toString()); }
  return "";
}

// === 5. OPENAI TTS (MODIFICADO: Retorna solo filename) ===
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
  
  // 👇 CAMBIO V3.6: Retornamos solo el nombre limpio
  return filename;
}

// === UTILS, SHEETS & EXPORT (LA MAGIA DE V3.6) ===

function cleanFilename(text) {
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + Utilities.getUuid().substring(0,4);
}

function extractFormData(e) {
  if (!e || !e.namedValues) return { palabra: "TEST_CLEAN", contexto: "Test context", modo: "Vocabulario General" };
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
  let sheet = ss.getSheetByName('Anki'); 
  if (!sheet) { sheet = ss.insertSheet('Anki'); }

  const headers = ['ID', 'Date', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Imported', 'Tags', 'Audio_Word', 'Image', 'Audio_Sentence'];
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === "" || firstCell !== 'ID') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#0d47a1').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function addToAnkiSheet(data) {
  const sheet = ensureAnkiSheet();
  let tagsClean = `${data.tag_mode} ${data.tags || ''}`.replace(/\s+/g, ' ').trim().replace('null', '');

  // Aquí se guardan los datos LIMPIOS (ej: word.mp3, imagen.png)
  sheet.appendRow([
    Utilities.getUuid().substring(0, 8),
    new Date().toLocaleDateString(),
    data.palabra,
    data.definicion,
    data.ejemplo,
    data.contexto,
    data.tipo,
    'NO',
    tagsClean,
    data.audioWord,     // Limpio
    data.image,         // Limpio
    data.audioSentence  // Limpio
  ]);
}

// ✅ EXPORTACIÓN INTELIGENTE: Reconstruye las etiquetas para Anki
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('Prepare New Words for Export', 'prepareAnkiExport')
    .addToUi();
}

function prepareAnkiExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Anki');
  if (!sourceSheet) { SpreadsheetApp.getUi().alert("No 'Anki' sheet found."); return; }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const statusIdx = headers.indexOf('Imported'); 
  if (statusIdx === -1) { SpreadsheetApp.getUi().alert("Column 'Imported' not found."); return; }

  const newWords = data.filter((row, index) => index > 0 && row[statusIdx] === 'NO');
  
  if (newWords.length === 0) {
    SpreadsheetApp.getUi().alert('No new words to export.');
    return;
  }

  let exportSheet = ss.getSheetByName('Anki_Export') || ss.insertSheet('Anki_Export');
  exportSheet.clear();
  
  const exportHeaders = ['ID', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Tags', 'Audio_Word', 'Image', 'Audio_Sentence'];
  exportSheet.getRange(1, 1, 1, exportHeaders.length).setValues([exportHeaders]).setFontWeight('bold');
  
  // Indices (V3.1): Word(2), Def(3), Ex(4), Ctx(5), Type(6), Tags(8), AudW(9), Img(10), AudS(11)
  const rowsToExport = newWords.map(r => {
    // 🪄 MAGIA V3.6: Aquí reconstruimos las etiquetas para Anki
    // Si la celda tiene texto, le ponemos el wrapper adecuado.
    const audioWordTag = r[9] ? `[sound:${r[9]}]` : "";
    const imageTag = r[10] ? `<img src="${r[10]}">` : "";
    const audioSentTag = r[11] ? `[sound:${r[11]}]` : "";

    return [
      r[0], // ID
      r[2], // Word
      r[3], // Definition
      r[4], // Example
      r[5], // Context
      r[6], // Type
      r[8], // Tags
      audioWordTag, // Audio Word (Con etiqueta)
      imageTag,     // Image (Con etiqueta)
      audioSentTag  // Audio Sent (Con etiqueta)
    ];
  });
  
  exportSheet.getRange(2, 1, rowsToExport.length, exportHeaders.length).setValues(rowsToExport);

  for (let i = 2; i <= sourceSheet.getLastRow(); i++) {
    if (sourceSheet.getRange(i, statusIdx + 1).getValue() === 'NO') {
      sourceSheet.getRange(i, statusIdx + 1).setValue('YES');
    }
  }

  exportSheet.activate();
  SpreadsheetApp.getUi().alert(`✅ Export listo. Hoja 'Anki' tiene nombres limpios, hoja 'Anki_Export' tiene etiquetas.`);
}

function testManualSubmission() {
  const mockEvent = {
    namedValues: {
      'Palabra o frase que quieres aprender': ['test_clean_v36'], 
      'Contexto u oración donde la viste (opcional)': ['Testing the V3.6 clean format.'],
      'Tipo de palabra (opcional)': ['noun'],
      'Modo de Estudio': ['Vocabulario General'] 
    }
  };
  console.log("🧪 Iniciando prueba V3.6...");
  processFormSubmission(mockEvent);
}
