// === CONFIGURATION ===
const CONFIG = {
   // Usa la misma Key de GCP si tiene permisos, o crea una nueva para Custom Search
  API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'), 

  //OpenAI models API to generate audio
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),

  //Drive folder code for Audios (replace with yours, otherwise it'll arrive to my folder)
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA', 

  //Drive folder code for Images (same as above)
  IMAGE_FOLDER_ID: '1TlEjDBQtyTYk0qBoalwkwCw3X8nA1Z4h'  
};

// === 1. INITIALIZATION ===
function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V5.0: Multimodal (Gemini + GPT-4o-mini + DALL-E 3).');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(ss).onFormSubmit().create();
}

// === HELPER: DRIVE SAVE ===
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
  console.log("🏁 INICIANDO PROCESO V5.0...");
  
  //extracción
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

  // === 2. CEREBRO (Gemini) ===
  let enriched;
  try {
    enriched = callGeminiAnalyst(wordData);
    console.log("✅ Gemini: Datos listos.");
  } catch (err) { console.error("❌ ERROR GEMINI:", err); return; }

  // === 3. AUDIO (OpenAI TTS) ===
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

  // === 4. IMAGEN (NUEVO FLUJO DALL-E) ===
  try {
    if (wordData.modo !== 'Solo Pronunciación') {
      console.log(`🎨 Generando prompt visual para: "${wordData.palabra}"...`);
      // GPT-4o-mini refina el query de Gemini para DALL-E
      const visualPrompt = generateVisualPrompt(enriched);
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.png`;
      enriched.image = callOpenAIDalle(visualPrompt, imgFilename);
    } else {
      enriched.image = "";
    }
  } catch (err) {
    console.error("⚠️ Error Imagen:", err.toString());
    enriched.image = ""; 
  }

  // === 5. GUARDAR ===
  try {
    addToAnkiSheet(enriched);
    console.log("🎉 ÉXITO TOTAL: Tarjeta guardada.");
  } catch (err) { console.error("❌ Error Sheets:", err); }
}

// === 6. GEMINI ANALYST (V4.5 RESTAURADA) ===
function callGeminiAnalyst(wordData) {
  const modelVersion = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;

  let promptText = "";
  
  // 🎤 MODO PRONUNCIACIÓN: FORZAMOS CLOZE
  if (wordData.modo === 'Solo Pronunciación') {
    promptText = `
      You are a linguistic engine. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      TASK: Create Pronunciation data for an Anki Cloze card.
      CRITICAL: You MUST use the cloze format {{c1::word}} in the 'example' field.
      JSON Schema:
      {
        "definition": "IPA transcription (e.g. /wɜːrd/).",
        "example": "Pronunciation tip for {{c1::${wordData.palabra}}}: [Insert tip regarding stress/linking]",
        "example_raw": "Pronunciation tip for ${wordData.palabra}: [Insert tip regarding stress/linking]",
        "type": "Part of speech",
        "image_query": null
      }
    `;
    // 📚 MODO VOCABULARIO GENERAL
  } else {
    promptText = `
      You are an expert IELTS vocabulary tutor.
      INPUT: Word: "${wordData.palabra}", Context: "${wordData.contexto}".
      TASK: Create Anki card JSON.
      RULES:
      1. Use Context to understand meaning.
      2. GENERATE A NEW "example" sentence.
      3. The "example" must be clear and use the word naturally.
      JSON Schema:
      {
        "definition": "Concise definition (max 15 words).",
        "example": "New sentence with Anki cloze: 'The {{c1::word}} ...'",
        "example_raw": "New sentence plain text.",
        "type": "Part of speech.",
        "image_query": "Optimized description for an image"
      }
    `;
  }

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.2 }
  };

  const options = {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) throw new Error("Gemini API Error: " + response.getContentText());

  let rawText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
  // Verifica si Gemini olvidó el {{c1::}} y lo arregla automáticamente
  rawText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
  const result = JSON.parse(rawText);

  // 🛡️ AUTO-CORRECCIÓN DE CLOZE (UNIVERSAL)
  // Si la palabra está en la frase, la envolvemos
  let finalExample = result.example;
  // Si no está, forzamos el formato al inicio para que Anki no falle
  if (finalExample && !finalExample.includes('{{c1::')) {
      const escapedWord = wordData.palabra.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp(`(${escapedWord})`, 'gi');
      if (regex.test(finalExample)) {
          finalExample = finalExample.replace(regex, '{{c1::$1}}');
      } else {
          finalExample = `Note on {{c1::${wordData.palabra}}}: ${finalExample}`;
      }
  }

  return {
    ...wordData, 
    definicion: result.definition,
    ejemplo: finalExample, 
    ejemplo_raw: result.example_raw, 
    tipo: result.type,
    tags: null,
    image_query: result.image_query,
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 7. NUEVAS FUNCIONES OPENAI (DIRECTOR + DALL-E) ===

function generateVisualPrompt(enriched) {
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: "You are a visual architect. Create a DALL-E 3 prompt. No text in image. Style: 3D render or photography." },
      { role: "user", content: `Word: ${enriched.palabra}. Meaning: ${enriched.definicion}. Example: ${enriched.ejemplo_raw}. Create a descriptive prompt.` }
    ]
  };
  const options = {
    method: "post", headers: { "Authorization": "Bearer " + CONFIG.OPENAI_API_KEY },
    contentType: "application/json", payload: JSON.stringify(payload)
  };
  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText()).choices[0].message.content.trim();
}

function callOpenAIDalle(prompt, filename) {
  const url = "https://api.openai.com/v1/images/generations";
  const payload = { model: "dall-e-3", prompt: prompt, n: 1, size: "1024x1024" };
  const options = {
    method: "post", headers: { "Authorization": "Bearer " + CONFIG.OPENAI_API_KEY },
    contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) throw new Error(json.error.message);

  const imageBlob = UrlFetchApp.fetch(json.data[0].url).getBlob().setName(filename);
  saveFileToDrive(imageBlob, filename, CONFIG.IMAGE_FOLDER_ID);
  return filename;
}

// === 8. TTS & UTILS (SIN CAMBIOS) ===
function callOpenAITTS(text, filename) {
  if (!text) return "";
  const url = "https://api.openai.com/v1/audio/speech";
  const payload = { model: "tts-1", input: text, voice: "nova" };
  const response = UrlFetchApp.fetch(url, {
    method: "post", headers: { "Authorization": "Bearer " + CONFIG.OPENAI_API_KEY },
    contentType: "application/json", payload: JSON.stringify(payload)
  });
  const blob = response.getBlob().setName(filename);
  saveFileToDrive(blob, filename, CONFIG.AUDIO_FOLDER_ID);
  return filename;
}
// === UTILS, SHEETS & EXPORT ===
function cleanFilename(text) {
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + Utilities.getUuid().substring(0,4);
}

function extractFormData(e) {
  if (!e || !e.namedValues) return { palabra: "TEST", contexto: "Test context", modo: "Vocabulario General" };
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
  if (!sheet) { sheet = ss.insertSheet('Anki'); }
  const headers = ['ID', 'Date', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Imported', 'Tags', 'Audio_Word', 'Image', 'Audio_Sentence'];
  const firstCell = sheet.getRange(1, 1).getValue();
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'ID') {
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
    Utilities.getUuid().substring(0, 8), new Date().toLocaleDateString(), data.palabra, data.definicion, data.ejemplo, data.contexto, data.tipo, 'NO', data.tag_mode, data.audioWord, data.image, data.audioSentence 
  ]);
}

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
  if (newWords.length === 0) { SpreadsheetApp.getUi().alert('No new words to export.'); return; }

  let exportSheet = ss.getSheetByName('Anki_Export') || ss.insertSheet('Anki_Export');
  exportSheet.clear();
  const exportHeaders = ['ID', 'Word', 'Definition', 'Example', 'Context', 'Type', 'Tags', 'Audio_Word', 'Image', 'Audio_Sentence'];
  exportSheet.getRange(1, 1, 1, exportHeaders.length).setValues([exportHeaders]).setFontWeight('bold');

  const rowsToExport = newWords.map(r => [
    r[0], r[2], r[3], r[4], r[5], r[6], r[8],
    r[9] ? `[sound:${r[9]}]` : "", 
    r[10] ? `<img src="${r[10]}">` : "", 
    r[11] ? `[sound:${r[11]}]` : ""
  ]);

  exportSheet.getRange(2, 1, rowsToExport.length, exportHeaders.length).setValues(rowsToExport);
  for (let i = 2; i <= sourceSheet.getLastRow(); i++) {
    if (sourceSheet.getRange(i, statusIdx + 1).getValue() === 'NO') sourceSheet.getRange(i, statusIdx + 1).setValue('YES');
  }
  exportSheet.activate();
  SpreadsheetApp.getUi().alert(`✅ Export listo.`);
}
