// === CONFIGURATION ===
const CONFIG = {
  // Usa la misma Key de GCP si tiene permisos, o crea una nueva para Custom Search
  API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'), 
  
  // Customized Search Engine API
  SEARCH_ENGINE_ID: PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID'),
  //OpenAI models API to generate audio
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  //Drive folder code for Audios
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA', 
  //Drive folder code for Images
  IMAGE_FOLDER_ID: '1TlEjDBQtyTYk0qBoalwkwCw3X8nA1Z4h'  
};

// === 1. INITIALIZATION ===
function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V4.0 Google Images: Búsqueda Real + Etiquetas Estrictas.');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(ss).onFormSubmit().create();
}

// === HELPER: ROBUST SAVE  ===
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
  console.log("🏁 INICIANDO PROCESO V4.0...");
  
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

  // 2. CEREBRO (Gemini)
  let enriched;
  try {
    enriched = callGeminiAnalyst(wordData);
    console.log("✅ Gemini: Datos listos.");
  } catch (err) { console.error("❌ ERROR GEMINI:", err); return; }

  // 3. AUDIO (OpenAI)
  try {
    // Generamos nombre limpio sin corchetes
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

  // 4. IMAGEN (Google Custom Search)
  // Buscamos una imagen real en internet
  try {
    if (wordData.modo !== 'Solo Pronunciación' && enriched.image_query) {
      console.log(`🔎 Buscando imagen: "${enriched.image_query}"...`);
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.jpg`;
      enriched.image = callGoogleImageSearch(enriched.image_query, imgFilename);
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

// === 3. GEMINI ANALYST (MODIFICADO PARA BÚSQUEDA) ===
function callGeminiAnalyst(wordData) {
  const modelVersion = 'gemini-2.5-flash'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;

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
        "image_query": null
      }
    `;
  } else {
    promptText = `
      You are a linguistic engine. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      Task: Create Anki card. Output JSON.
      
      CRITICAL FOR IMAGE_QUERY:
      - Instead of describing an image, provide the BEST GOOGLE SEARCH QUERY to find a visual representation.
      - Prefer "vector", "icon", or "illustration" keywords in the query to get clean images.
      - Example: If word is "Run", query should be "person running flat vector icon".
      
      JSON Schema:
      {
        "definition": "Concise definition (max 15 words).",
        "example": "Sentence with Anki cloze: 'The {{c1::word}} ...'",
        "example_raw": "Same sentence plain text for Audio TTS.",
        "type": "Part of speech.",
        "image_query": "Optimized Google Images search query (e.g. 'word vector icon')"
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
    tags: null,
    image_query: result.image_query, // Ahora es una query de búsqueda, no un prompt
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 4. GOOGLE CUSTOM SEARCH (NUEVO) ===
function callGoogleImageSearch(query, filename) {
  if (!CONFIG.SEARCH_ENGINE_ID) {
    console.warn("⚠️ No Search Engine ID configured.");
    return "";
  }

  // Construimos la URL de la API
  // searchType=image: Solo busca imágenes
  // num=1: Solo queremos 1 resultado
  // fileType=jpg|png: Preferimos formatos estándar
  // safe=active: Filtro SafeSearch activado
  const apiUrl = `https://www.googleapis.com/customsearch/v1?key=${CONFIG.API_KEY}&cx=${CONFIG.SEARCH_ENGINE_ID}&q=${encodeURIComponent(query)}&searchType=image&num=1&safe=active&fileType=jpg,png`;

  try {
    const response = UrlFetchApp.fetch(apiUrl, {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) {
      console.warn("⚠️ Error en búsqueda Google:", response.getContentText());
      return "";
    }

    const json = JSON.parse(response.getContentText());
    
    // Verificamos si hay resultados
    if (!json.items || json.items.length === 0) {
      console.warn("⚠️ No se encontraron imágenes para:", query);
      return "";
    }

    // Tomamos la URL de la primera imagen
    const imageUrl = json.items[0].link;
    console.log(`🖼️ Imagen encontrada: ${imageUrl}`);

    // Descargamos la imagen
    const imageResponse = UrlFetchApp.fetch(imageUrl, {muteHttpExceptions: true});
    if (imageResponse.getResponseCode() !== 200) {
      console.warn("⚠️ No se pudo descargar la imagen remota.");
      return "";
    }

    // Guardamos en Drive
    const blob = imageResponse.getBlob().setName(filename);
    saveFileToDrive(blob, filename, CONFIG.IMAGE_FOLDER_ID);
    return filename;

  } catch (e) {
    console.error("Excepción en Búsqueda de Imagen:", e.toString());
    return "";
  }
}

// === 5. OPENAI TTS (Sin cambios) ===
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
  return filename;
}

// === UTILS, SHEETS & EXPORT ===

function cleanFilename(text) {
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + Utilities.getUuid().substring(0,4);
}

function extractFormData(e) {
  if (!e || !e.namedValues) return { palabra: "TEST_GOOGLE", contexto: "Test context", modo: "Vocabulario General" };
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
  const finalTag = data.tag_mode; 

  sheet.appendRow([
    Utilities.getUuid().substring(0, 8),
    new Date().toLocaleDateString(),
    data.palabra,
    data.definicion,
    data.ejemplo,
    data.contexto,
    data.tipo,
    'NO',
    finalTag,
    data.audioWord,     // Guardamos SOLO el nombre del archivo
    data.image,         // Guardamos SOLO el nombre del archivo
    data.audioSentence  // Guardamos SOLO el nombre del archivo
  ]);
}

// ✅ EXPORTACIÓN (V3.7 Restaurada - Agrega corchetes AQUÍ)
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
  
  const rowsToExport = newWords.map(r => {
    // 🛑 AQUÍ SE AGREGAN LOS CORCHETES [sound:...]
    // Como en la hoja 'Anki' solo guardamos el nombre limpio, esto funcionará perfecto.
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
      audioWordTag,
      imageTag,
      audioSentTag
    ];
  });
  
  exportSheet.getRange(2, 1, rowsToExport.length, exportHeaders.length).setValues(rowsToExport);

  for (let i = 2; i <= sourceSheet.getLastRow(); i++) {
    if (sourceSheet.getRange(i, statusIdx + 1).getValue() === 'NO') {
      sourceSheet.getRange(i, statusIdx + 1).setValue('YES');
    }
  }

  exportSheet.activate();
  SpreadsheetApp.getUi().alert(`✅ Export listo.`);
}

function testManualSubmission() {
  const mockEvent = {
    namedValues: {
      'Palabra o frase que quieres aprender': ['laptop'], 
      'Contexto u oración donde la viste (opcional)': [''],
      'Tipo de palabra (opcional)': [''],
      'Modo de Estudio': ['Vocabulario General'] 
    }
  };
  console.log("🧪 Iniciando prueba V4.0 (Google Images)...");
  processFormSubmission(mockEvent);
}
