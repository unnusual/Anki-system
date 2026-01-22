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
    .addItem('♾️ REGENERATE ALL (Auto-Run)', 'startBatchProcess')
    .addItem('🛑 STOP Auto-Run', 'stopBatchProcess')
    .addToUi();
}

function initializeSystem() {
  ensureAnkiSheet();
  setupTrigger();
  console.log('🚀 Sistema V6.2: Juez Estricto (Reasoning).');
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteAllTriggers('batchRegenerateMedia');
  deleteAllTriggers('processFormSubmission');
  ScriptApp.newTrigger('processFormSubmission').forSpreadsheet(ss).onFormSubmit().create();
}

function deleteAllTriggers(funcName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// === HELPER: GUARDADO ===
function saveFileToDrive(blob, filename, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const existing = folder.getFilesByName(filename);
    if (existing.hasNext()) return true; 
    const fileMetadata = { title: filename, parents: [{id: folderId}], mimeType: blob.getContentType() };
    Drive.Files.insert(fileMetadata, blob);
    return true; 
  } catch (e) {
    console.error(`❌ Error guardando ${filename}: ${e.message}`);
    return false;
  }
}

// === 🧠 PASO 1: QUERY (GEMINI 2.5) ===
function getSmartImageQuery(word, cleanSentence, context) {
  const modelVersion = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;
  const baseContext = cleanSentence || context || word;

  const promptText = `
    You are a visual search expert.
    Task: Create ONE Google Images search query for: "${word}".
    Context: "${baseContext}".

    RULES:
    1. Describe the VISUAL SCENE accurately.
    2. "Bury" -> "person digging hole with shovel in dirt photography" OR "dog burying bone".
    3. AVOID ABSTRACT TERMS. Be descriptive of the physical action.
    4. FORBIDDEN: text, slides, album covers, cartoons.
    OUTPUT: ONLY the search query string.
  `;

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" }
    ],
    generationConfig: { temperature: 0.3, maxOutputTokens: 50 }
  };

  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return `${word} action photography -text`;
    const json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0].content) return `${word} action photography -text`;
    let query = json.candidates[0].content.parts[0].text.trim();
    query = query.replace(/^"|"$/g, '').replace(/\n/g, ' ');
    return query.length < 3 ? `${word} action photography -text` : query;
  } catch (e) { return `${word} action photography -text`; }
}

// === 👁️ PASO 2: EL JUEZ ESTRICTO (VALIDACIÓN) ===
function verifyImageContent(imageBlob, word, context) {
  // 🔄 CAMBIO: Usamos 'gemini-2.5-pro' que sabemos que sí tienes activo
  const modelVersion = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;
  
  const base64Image = Utilities.base64Encode(imageBlob.getBytes());
  const mimeType = imageBlob.getContentType();

  // 🔥 PROMPT ESTRICTO CON RAZONAMIENTO
  const promptText = `
    Act as a strict Quality Assurance Image Validator.
    Target Concept: "${word}"
    Context: "${context}"

    Task: Analyze the image and determine if it matches the concept perfectly.
    
    1. First, describe briefly what you see in the image.
    2. Then, compare it to the target concept.
    3. REJECT if: 
       - It is completely unrelated.
       - It contains text overlay, slides, or charts.
       - It is an album cover or logo.
       - It is confusing or low quality.
    
    Output JSON format ONLY:
    {
      "description": "short description of image",
      "reason": "why it matches or fails",
      "verdict": "PASS" or "FAIL"
    }
  `;

  const payload = {
    contents: [{ parts: [{ text: promptText }, { inlineData: { mimeType: mimeType, data: base64Image } }] }],
    safetySettings: [
        { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" }, 
        { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" }
    ],
    generationConfig: { responseMimeType: "application/json", temperature: 0.0 }
  };

  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
        // Ahora sí podemos ver el error real si vuelve a pasar
        console.warn(`      ⚠️ Error API Vision (${response.getResponseCode()}): ${response.getContentText()}`);
        return false; 
    }
    
    const json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0].content) {
       console.warn("      ⚠️ Gemini devolvió una respuesta vacía.");
       return false;
    }

    const content = json.candidates[0].content.parts[0].text;
    const result = JSON.parse(content);
    
    console.log(`      🤖 Juez: ${result.verdict} | Razón: ${result.reason}`);
    
    return result.verdict === "PASS";

  } catch (e) {
    console.warn(`      ⚠️ Excepción Juez: ${e.toString()}. Rechazando.`);
    return false; 
  }
}

// === PASO 3: BÚSQUEDA Y SELECCIÓN ===
function callGoogleImageSearchAndValidate(query, filename, word, context) {
  if (!CONFIG.SEARCH_ENGINE_ID) return "";
  const cleanQuery = query + " -text -slide -lyrics -album -screenshot";
  // Pedimos 3 candidatos
  const apiUrl = `https://www.googleapis.com/customsearch/v1?key=${CONFIG.API_KEY}&cx=${CONFIG.SEARCH_ENGINE_ID}&q=${encodeURIComponent(cleanQuery)}&searchType=image&num=3&safe=active`;

  try {
    const response = UrlFetchApp.fetch(apiUrl, {muteHttpExceptions: true});
    if (response.getResponseCode() !== 200) return "";
    const json = JSON.parse(response.getContentText());
    if (!json.items || json.items.length === 0) return "";

    for (let i = 0; i < json.items.length; i++) {
      const item = json.items[i];
      const imageUrl = item.link;
      console.log(`   🔎 Evaluando #${i + 1}...`);

      try {
        const imageResponse = UrlFetchApp.fetch(imageUrl, {muteHttpExceptions: true});
        if (imageResponse.getResponseCode() !== 200) continue; 
        const blob = imageResponse.getBlob();
        
        // Validación Estricta
        const isValid = verifyImageContent(blob, word, context);
        
        if (isValid) {
          console.log("   ✅ ¡Aprobada!");
          let extension = ".jpg";
          if (imageUrl.toLowerCase().endsWith(".png")) extension = ".png";
          const finalFilename = filename.replace(/\.(jpg|png)$/, "") + extension;
          blob.setName(finalFilename);
          saveFileToDrive(blob, finalFilename, CONFIG.IMAGE_FOLDER_ID);
          return finalFilename; 
        } else {
            console.log("   ❌ Rechazada. Siguiente...");
        }
      } catch (err) {}
    }
    console.warn("   ⚠️ Ninguna imagen pasó el filtro estricto. Dejando vacío.");
    return ""; // Retorna vacío si ninguna sirve
  } catch (e) { return ""; }
}

// === ♾️ INFINITE RUNNER ===
function startBatchProcess() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('♾️ Infinite Runner', 'Iniciando modo estricto V6.2.\nSe ejecutará cada 5 minutos.\n¿Comenzar?', ui.ButtonSet.YES_NO);
  if (result === ui.Button.YES) { batchRegenerateMedia(); }
}

function stopBatchProcess() {
  deleteAllTriggers('batchRegenerateMedia');
  SpreadsheetApp.getUi().alert('🛑 Detenido.');
}

function batchRegenerateMedia() {
  deleteAllTriggers('batchRegenerateMedia');
  const startTime = new Date().getTime(); 
  const TIME_LIMIT = 270 * 1000; // 4.5 Minutos

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAnkiSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  const WORD_IDX = 2; const EXAMPLE_IDX = 4; const CONTEXT_IDX = 5; 
  const TAG_MODE_IDX = 8; const AUDIO_WORD_IDX = 9; const IMAGE_IDX = 10; const AUDIO_SENT_IDX = 11;

  let processedCount = 0;
  let finishedAll = true;

  console.log('🔄 INICIANDO LOTE V6.2 (Estricto)...');

  for (let i = 1; i < data.length; i++) {
    if (new Date().getTime() - startTime > TIME_LIMIT) {
      console.warn("⏳ Tiempo límite. Reprogramando...");
      ScriptApp.newTrigger('batchRegenerateMedia').timeBased().after(1000 * 30).create();
      finishedAll = false;
      break; 
    }

    const rowNumber = i + 1;
    const word = data[i][WORD_IDX];
    if (!word) continue;

    const tagMode = data[i][TAG_MODE_IDX];
    const hasWordAudio = data[i][AUDIO_WORD_IDX] !== "";
    const hasImage = data[i][IMAGE_IDX] !== "";
    const hasSentAudio = data[i][AUDIO_SENT_IDX] !== "";

    let isComplete = (tagMode === 'pronunciation') ? (hasWordAudio && hasSentAudio) : (hasWordAudio && hasImage && hasSentAudio);
    if (isComplete) continue; 

    console.log(`🔹 [Fila ${rowNumber}] Procesando: "${word}"`);

    const baseFilename = cleanFilename(word);
    const newWordFile = `word_${baseFilename}.mp3`;
    const newSentFile = `sent_${baseFilename}.mp3`;
    const newImgFile = `img_${baseFilename}.jpg`;

    // 1. Audio Palabra
    if (!hasWordAudio) {
      if(callOpenAITTS(word, newWordFile)) sheet.getRange(rowNumber, AUDIO_WORD_IDX + 1).setValue(newWordFile);
    }

    // Preparar Texto
    let rawExample = data[i][EXAMPLE_IDX];
    let cleanSentenceText = "";
    if (rawExample && rawExample.toString().trim() !== "") {
       let str = rawExample.toString();
       if(str.includes('{{c1::')) cleanSentenceText = str.replace(/\{\{c\d+::(.*?)\}\}/g, '$1');
       else cleanSentenceText = str; 
    }

    // 2. Audio Frase
    if (!hasSentAudio && cleanSentenceText) {
      if(callOpenAITTS(cleanSentenceText, newSentFile)) sheet.getRange(rowNumber, AUDIO_SENT_IDX + 1).setValue(newSentFile);
    }

    // 3. Imagen Validada
    if (tagMode !== 'pronunciation' && !hasImage) {
      const context = data[i][CONTEXT_IDX];
      const smartQuery = getSmartImageQuery(word, cleanSentenceText, context);
      console.log(`   🎨 Query: "${smartQuery}"`);
      const imgResult = callGoogleImageSearchAndValidate(smartQuery, newImgFile, word, cleanSentenceText || context);
      if (imgResult) {
          sheet.getRange(rowNumber, IMAGE_IDX + 1).setValue(newImgFile);
      } else {
          // Si no hay imagen válida, marcamos temporalmente como "FAILED" o lo dejamos vacío
          // Dejarlo vacío permite reintentar en el futuro.
          console.warn(`   ⚠️ Fila ${rowNumber}: No se encontró imagen válida para "${word}".`);
      }
    }
    
    processedCount++;
    SpreadsheetApp.flush(); 
  }

  if (finishedAll) {
    try { ss.toast("🎉 Proceso finalizado.", "Fin"); } catch(e){}
  } else {
    try { ss.toast(`⏳ Pausa por tiempo. Reiniciando en 30s...`, "Continuando..."); } catch(e){}
  }
}

// === MAIN PROCESSOR (Formulario) ===
function processFormSubmission(e) {
  console.log("🏁 INICIANDO PROCESO INDIVIDUAL V6.2...");
  let wordData;
  try { wordData = extractFormData(e); if (!wordData.palabra) return; } catch (err) { return; }
  const sheet = ensureAnkiSheet();
  const existingWords = sheet.getRange("C:C").getValues().flat().filter(c => c !== "").map(w => w.toString().toLowerCase());
  if (existingWords.includes(wordData.palabra.toLowerCase())) return;

  let enriched;
  try { enriched = callGeminiAnalyst(wordData); } catch (err) { console.error(err); return; }

  const wFile = `word_${cleanFilename(wordData.palabra)}.mp3`;
  enriched.audioWord = callOpenAITTS(wordData.palabra, wFile);
  
  if (enriched.ejemplo_raw && wordData.modo !== 'Solo Pronunciación') {
     const sFile = `sent_${cleanFilename(wordData.palabra)}.mp3`;
     enriched.audioSentence = callOpenAITTS(enriched.ejemplo_raw, sFile);
  } else { enriched.audioSentence = ""; }

  if (wordData.modo !== 'Solo Pronunciación' && enriched.tag_mode !== 'pronunciation' && enriched.image_query) {
     const iFile = `img_${cleanFilename(wordData.palabra)}.jpg`;
     enriched.image = callGoogleImageSearchAndValidate(enriched.image_query, iFile, wordData.palabra, enriched.ejemplo_raw || wordData.contexto);
  } else { enriched.image = ""; }

  addToAnkiSheet(enriched);
}

// === FUNCIONES SOPORTE (TTS, Utils) ===
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
  } catch(e) { return ""; }
}

function callGeminiAnalyst(wordData) {
  const modelVersion = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;
  const promptText = `Analyze: "${wordData.palabra}". Context: "${wordData.contexto}". Create Anki card JSON. Image Query: Visual scene description. JSON: {definition, example, example_raw, type, image_query}`;
  const payload = { contents: [{ parts: [{ text: promptText }] }], safetySettings: [{category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE"}, {category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE"}], generationConfig: { responseMimeType: "application/json", temperature: 0.2 } };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text);
  return { ...wordData, definicion: result.definition, ejemplo: result.example, ejemplo_raw: result.example_raw, tipo: result.type, image_query: result.image_query, tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab' };
}

function cleanFilename(text) { return text.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 15) + "_" + new Date().getTime().toString().substring(8); }
function extractFormData(e) { if(!e) return {palabra:"TEST"}; const v=e.namedValues; return { palabra: v['Palabra o frase que quieres aprender']?v['Palabra o frase que quieres aprender'][0].trim():'', contexto: v['Contexto u oración donde la viste (opcional)']?v['Contexto u oración donde la viste (opcional)'][0].trim():'', modo: v['Modo de Estudio']?v['Modo de Estudio'][0].trim():'Vocabulario General' }; }
function ensureAnkiSheet() { const ss=SpreadsheetApp.getActiveSpreadsheet(); let s=ss.getSheetByName('Anki'); if(!s) s=ss.insertSheet('Anki'); return s; }
function addToAnkiSheet(d) { const s=ensureAnkiSheet(); s.appendRow([Utilities.getUuid().substring(0,8), new Date().toLocaleDateString(), d.palabra, d.definicion, d.ejemplo, d.contexto, d.tipo, 'NO', d.tag_mode, d.audioWord, d.image, d.audioSentence]); }
function prepareAnkiExport() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const s = ss.getSheetByName('Anki'); const d = s.getDataRange().getValues(); const h = d[0]; const idx = h.indexOf('Imported');
  const nw = d.filter((r, i) => i > 0 && r[idx] === 'NO'); if (!nw.length) { SpreadsheetApp.getUi().alert('No new words.'); return; }
  let es = ss.getSheetByName('Anki_Export') || ss.insertSheet('Anki_Export'); es.clear();
  es.getRange(1, 1, 1, 10).setValues([['ID','Word','Definition','Example','Context','Type','Tags','Audio_Word','Image','Audio_Sentence']]);
  const rows = nw.map(r => [r[0], r[2], r[3], r[4], r[5], r[6], r[8], r[9]?`[sound:${r[9]}]`:'', (r[10]&&r[10].match(/\.(jpg|png)$/))?`<img src="${r[10]}">`:'', r[11]?`[sound:${r[11]}]`:'']);
  es.getRange(2, 1, rows.length, 10).setValues(rows);
  for(let i=2; i<=s.getLastRow(); i++) { if(s.getRange(i, idx+1).getValue()==='NO') s.getRange(i, idx+1).setValue('YES'); }
  SpreadsheetApp.getUi().alert('Export ready.');
}
