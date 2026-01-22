// === CONFIGURATION ===
const CONFIG = {
  API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  SEARCH_ENGINE_ID: PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID'),
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  
  AUDIO_FOLDER_ID: '1HKTOv1SwgP4HYmKwY6O7A0XQyvr7ihrA', 
  IMAGE_FOLDER_ID: '1TlEjDBQtyTYk0qBoalwkwCw3X8nA1Z4h'  
};

// === 1. INITIALIZATION & MENU ===
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ Anki Tools')
    .addItem('✅ Prepare New Words for Export', 'prepareAnkiExport')
    .addSeparator()
    .addItem('♻️ REGENERATE MEDIA (Cycle Batch)', 'batchRegenerateMedia')
    .addToUi();
}

function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V5.4: Safety OFF + Robustez.');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(ss).onFormSubmit().create();
}

// === HELPER: GUARDADO ROBUSTO ===
function saveFileToDrive(blob, filename, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const existing = folder.getFilesByName(filename);
    if (existing.hasNext()) return true; 
    
    const fileMetadata = {
      title: filename,
      parents: [{id: folderId}],
      mimeType: blob.getContentType()
    };
    Drive.Files.insert(fileMetadata, blob);
    return true; 
  } catch (e) {
    console.error(`❌ Error guardando archivo ${filename}: ${e.message}`);
    return false;
  }
}

// === 🧠 SMART IMAGE QUERY V5.4 (SAFETY OFF) ===
function getSmartImageQuery(word, cleanSentence, context) {
  const modelVersion = 'gemini-2.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;

  const baseContext = cleanSentence || context || word;

  const promptText = `
    You are a visual director.
    Task: Create a Google Images search query for: "${word}".
    Context/Scene: "${baseContext}"

    RULES:
    1. Visualize the human emotion, action, or physical scene described.
    2. "Hook up" -> "connecting cables close up" OR "people meeting social gathering" (depending on context).
    3. "Coax" -> "person gently persuading another smiling photography".
    4. NO TEXT/SLIDES.
    5. OUTPUT: ONLY the search query string.
  `;

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    // 🔥 CONFIGURACIÓN CRÍTICA: Desactivar filtros de seguridad
    safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
    ],
    generationConfig: { temperature: 0.3, maxOutputTokens: 50 }
  };

  const options = {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    
    // DEBUG: Ver si falló la API
    if (response.getResponseCode() !== 200) {
        console.warn(`⚠️ API Error (${response.getResponseCode()}): ${response.getContentText()}`);
        return `${word} photography -text`;
    }
    
    const json = JSON.parse(response.getContentText());
    
    // Verificar si hay candidatos válidos
    if (!json.candidates || json.candidates.length === 0 || !json.candidates[0].content) {
        console.warn(`⚠️ Gemini bloqueó la respuesta (Safety?) para: ${word}`);
        return `${word} photography -text`;
    }

    let query = json.candidates[0].content.parts[0].text.trim();
    query = query.replace(/^"|"$/g, '').replace(/\n/g, ' ');

    // VALIDACIÓN: Si responde basura como "z" o vacío
    if (query.length < 3) {
        console.warn(`⚠️ Respuesta inválida ("${query}"), usando fallback.`);
        return `${word} photography -text`;
    }

    return query;
  } catch (e) {
    console.warn(`⚠️ Excepción Smart Query: ${e.toString()}`);
    return `${word} photography -text`;
  }
}

// === REGENERADOR POR CICLOS ===
function batchRegenerateMedia() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!CONFIG.SEARCH_ENGINE_ID) {
    ui.alert('❌ ERROR', 'Falta SEARCH_ENGINE_ID en Propiedades.', ui.ButtonSet.OK);
    return;
  }

  const BATCH_LIMIT = 12; 

  const sheet = ensureAnkiSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  const WORD_IDX = 2;       
  const EXAMPLE_IDX = 4;   
  const CONTEXT_IDX = 5;
  const TAG_MODE_IDX = 8;
  const AUDIO_WORD_IDX = 9; 
  const IMAGE_IDX = 10;     
  const AUDIO_SENT_IDX = 11;

  let processedCount = 0;
  let skippedCount = 0;

  ss.toast(`Iniciando V5.4 (Safety OFF)...`, '⏳ Procesando');
  console.log('🔄 INICIANDO CICLO V5.4...');

  for (let i = 1; i < data.length; i++) {
    if (processedCount >= BATCH_LIMIT) break;

    const rowNumber = i + 1;
    const word = data[i][WORD_IDX];
    const tagMode = data[i][TAG_MODE_IDX];
    
    const hasWordAudio = data[i][AUDIO_WORD_IDX] !== "";
    const hasImage = data[i][IMAGE_IDX] !== "";
    const hasSentAudio = data[i][AUDIO_SENT_IDX] !== "";

    let isComplete = false;
    if (tagMode === 'pronunciation') {
        isComplete = hasWordAudio && hasSentAudio;
    } else {
        isComplete = hasWordAudio && hasImage && hasSentAudio;
    }

    if (isComplete) {
      skippedCount++;
      continue;
    }

    if (!word) continue;

    console.log(`🔹 [Fila ${rowNumber}] Procesando: "${word}"`);

    const baseFilename = cleanFilename(word);
    const newWordFile = `word_${baseFilename}.mp3`;
    const newSentFile = `sent_${baseFilename}.mp3`;
    const newImgFile = `img_${baseFilename}.jpg`;

    // 1. AUDIO PALABRA
    if (!hasWordAudio) {
      const wordResult = callOpenAITTS(word, newWordFile);
      if (wordResult) sheet.getRange(rowNumber, AUDIO_WORD_IDX + 1).setValue(newWordFile);
    }

    // 2. TEXTO LIMPIO
    let rawExample = data[i][EXAMPLE_IDX];
    let cleanSentenceText = "";
    
    if (rawExample && rawExample.toString().trim() !== "") {
      let finalExampleStr = rawExample.toString();
      if (finalExampleStr.includes('{{c1::')) {
        cleanSentenceText = finalExampleStr.replace(/\{\{c\d+::(.*?)\}\}/g, '$1');
      } else {
        const escapedWord = word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); 
        const regex = new RegExp(`(${escapedWord})`, 'gi'); 
        if (regex.test(finalExampleStr)) {
          const fixedExample = finalExampleStr.replace(regex, '{{c1::$1}}');
          sheet.getRange(rowNumber, EXAMPLE_IDX + 1).setValue(fixedExample);
          cleanSentenceText = finalExampleStr; 
        } else {
          cleanSentenceText = finalExampleStr;
        }
      }
    }

    // 3. AUDIO FRASE
    if (!hasSentAudio && cleanSentenceText) {
      let textToSpeak = cleanSentenceText.replace(/\[\.\.\.\]/g, "something").trim();
      const sentResult = callOpenAITTS(textToSpeak, newSentFile);
      if (sentResult) sheet.getRange(rowNumber, AUDIO_SENT_IDX + 1).setValue(newSentFile);
    }

    // 4. IMAGEN (Con Safety OFF)
    if (tagMode !== 'pronunciation' && !hasImage) {
      const context = data[i][CONTEXT_IDX];
      const smartQuery = getSmartImageQuery(word, cleanSentenceText, context);
      console.log(`   🎨 Query: "${smartQuery}"`);
      
      const imgResult = callGoogleImageSearch(smartQuery, newImgFile);
      if (imgResult) sheet.getRange(rowNumber, IMAGE_IDX + 1).setValue(newImgFile);
    }
    
    processedCount++;
    SpreadsheetApp.flush(); 
    Utilities.sleep(1500); 
  }

  console.log(`✅ CICLO TERMINADO. Procesados: ${processedCount}.`);
  
  if (processedCount >= BATCH_LIMIT) {
    ui.alert('⏳ Ciclo completado', `Se procesaron ${processedCount} filas.\nEjecuta de nuevo.`, ui.ButtonSet.OK);
  } else {
    ui.alert('🎉 Todo listo', `Proceso finalizado.`, ui.ButtonSet.OK);
  }
}

// === 2. MAIN PROCESSOR (Formulario) ===
function processFormSubmission(e) {
  console.log("🏁 INICIANDO PROCESO V5.4...");
  
  let wordData;
  try {
    wordData = extractFormData(e);
    if (!wordData.palabra) { console.warn("⚠️ No se detectó palabra."); return; }
  } catch (err) { console.error("❌ Error Data:", err); return; }

  const sheet = ensureAnkiSheet();
  const existingWords = sheet.getRange("C:C").getValues().flat()
    .filter(cell => cell !== "") 
    .map(w => w.toString().toLowerCase());

  if (existingWords.includes(wordData.palabra.toLowerCase())) {
    console.warn(`⏭️ DUPLICADO: "${wordData.palabra}".`);
    return;
  }

  let enriched;
  try {
    enriched = callGeminiAnalyst(wordData);
  } catch (err) { console.error("❌ ERROR GEMINI:", err); return; }

  try {
    const wordFilename = `word_${cleanFilename(wordData.palabra)}.mp3`;
    enriched.audioWord = callOpenAITTS(wordData.palabra, wordFilename);
    
    if (enriched.ejemplo_raw && wordData.modo !== 'Solo Pronunciación') {
       const sentenceFilename = `sent_${cleanFilename(wordData.palabra)}.mp3`;
       enriched.audioSentence = callOpenAITTS(enriched.ejemplo_raw, sentenceFilename);
    } else {
       enriched.audioSentence = "";
    }
  } catch (err) {
    console.error("⚠️ Error Audio:", err);
    enriched.audioWord = ""; enriched.audioSentence = "";
  }

  try {
    if (wordData.modo !== 'Solo Pronunciación' && enriched.tag_mode !== 'pronunciation' && enriched.image_query) {
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.jpg`;
      enriched.image = callGoogleImageSearch(enriched.image_query, imgFilename);
    } else {
      enriched.image = "";
    }
  } catch (err) {
    console.error("⚠️ Error Imagen:", err.toString());
    enriched.image = ""; 
  }

  try {
    addToAnkiSheet(enriched);
  } catch (err) { console.error("❌ Error Sheets:", err); }
}

// === 3. GEMINI ANALYST ===
function callGeminiAnalyst(wordData) {
  const modelVersion = 'gemini-2.5-pro'; 
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
      Task: Create Anki card JSON.
      CRITICAL: Create a Google Images search query describing the SCENE in the example sentence.
      - FOCUS ON HUMAN EMOTION/ACTION for abstract words.
      - FORBIDDEN: "text", "slide", "song", "album cover".
      
      JSON Schema:
      {
        "definition": "Concise definition (max 15 words).",
        "example": "Sentence with Anki cloze: 'The {{c1::word}} ...'",
        "example_raw": "Same sentence plain text for Audio TTS.",
        "type": "Part of speech.",
        "image_query": "Optimized Google Images search query"
      }
    `;
  }

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" }
    ],
    generationConfig: { responseMimeType: "application/json", temperature: 0.15 }
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
    image_query: result.image_query,
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 4. GOOGLE SEARCH ===
function callGoogleImageSearch(query, filename) {
  if (!CONFIG.SEARCH_ENGINE_ID) return "";
  const cleanQuery = query + " -text -slide -lyrics -album -screenshot";
  const apiUrl = `https://www.googleapis.com/customsearch/v1?key=${CONFIG.API_KEY}&cx=${CONFIG.SEARCH_ENGINE_ID}&q=${encodeURIComponent(cleanQuery)}&searchType=image&num=1&safe=active`;

  try {
    const response = UrlFetchApp.fetch(apiUrl, {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) return "";
    const json = JSON.parse(response.getContentText());
    if (!json.items || json.items.length === 0) return "";
    const imageUrl = json.items[0].link;
    let extension = ".jpg";
    if (imageUrl.toLowerCase().endsWith(".png")) extension = ".png";
    const finalFilename = filename.replace(/\.(jpg|png)$/, "") + extension;
    const imageResponse = UrlFetchApp.fetch(imageUrl, {muteHttpExceptions: true});
    if (imageResponse.getResponseCode() !== 200) return "";
    const blob = imageResponse.getBlob().setName(finalFilename);
    saveFileToDrive(blob, finalFilename, CONFIG.IMAGE_FOLDER_ID);
    return finalFilename;
  } catch (e) {
    console.error("Excepción Búsqueda Imagen:", e.toString());
    return "";
  }
}

// === 5. OPENAI TTS ===
function callOpenAITTS(text, filename) {
  if (!text) return "";
  const url = "https://api.openai.com/v1/audio/speech";
  const payload = { model: "tts-1", input: text, voice: "nova", response_format: "mp3" };
  const options = { method: "post", headers: { "Authorization": "Bearer " + CONFIG.OPENAI_API_KEY }, contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return "";
    const blob = response.getBlob().setName(filename);
    saveFileToDrive(blob, filename, CONFIG.AUDIO_FOLDER_ID);
    return filename;
  } catch(e) { console.error("Excepción TTS:", e.toString()); return ""; }
}

// === UTILS ===
function cleanFilename(text) {
  const timestamp = new Date().getTime().toString().substring(8); 
  return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + timestamp;
}
function extractFormData(e) {
  if (!e || !e.namedValues) return { palabra: "TEST_V54", contexto: "Debug context", modo: "Vocabulario General" };
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
  if (firstCell === "" || firstCell !== 'ID') { sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#0d47a1').setFontColor('white'); sheet.setFrozenRows(1); }
  return sheet;
}
function addToAnkiSheet(data) {
  const sheet = ensureAnkiSheet();
  const finalTag = data.tag_mode; 
  sheet.appendRow([Utilities.getUuid().substring(0, 8), new Date().toLocaleDateString(), data.palabra, data.definicion, data.ejemplo, data.contexto, data.tipo, 'NO', finalTag, data.audioWord, data.image, data.audioSentence]);
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
  const rowsToExport = newWords.map(r => {
    const audioWordTag = r[9] ? `[sound:${r[9]}]` : "";
    let imageTag = "";
    if (r[10]) { if (r[10].endsWith('.jpg') || r[10].endsWith('.png')) { imageTag = `<img src="${r[10]}">`; } }
    const audioSentTag = r[11] ? `[sound:${r[11]}]` : "";
    return [r[0], r[2], r[3], r[4], r[5], r[6], r[8], audioWordTag, imageTag, audioSentTag];
  });
  exportSheet.getRange(2, 1, rowsToExport.length, exportHeaders.length).setValues(rowsToExport);
  for (let i = 2; i <= sourceSheet.getLastRow(); i++) { if (sourceSheet.getRange(i, statusIdx + 1).getValue() === 'NO') { sourceSheet.getRange(i, statusIdx + 1).setValue('YES'); } }
  exportSheet.activate();
  SpreadsheetApp.getUi().alert(`✅ Export listo.`);
}
