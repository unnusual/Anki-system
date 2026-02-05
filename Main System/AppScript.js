// === CONFIGURATION ===
const CONFIG = {
   // Gemini API to generate content
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
  console.log('🚀 V5.0 System: Multimodal (Gemini + GPT-4o-mini + DALL-E 3).');
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
    console.error(`❌ Error saving archive ${filename}: ${e.message}`);
    throw e;
  }
}

// === 2. MAIN PROCESSOR ===
function processFormSubmission(e) {
  console.log("🏁 INITIALIZING PROCESS V5.0...");
  
  // Extraction
  const wordData = extractFormData(e);
  try {
    if (!wordData.palabra) { console.warn("⚠️ Word not found"); return; }
    console.log(`📌 Processing: "${wordData.palabra}"`);
  } catch (err) { console.error("❌ Error Data:", err); return; }

  // --- 🛡️ SMART DEDUPLICATION LOGIC ---
  const sheet = ensureAnkiSheet();
  const allData = sheet.getDataRange().getValues();

  const existingRow = allData.find(row => 
    row[2] && row[2].toString().toLowerCase() === wordData.palabra.toLowerCase()
  );

  let isPolysemy = false; // Flag for knowing if it's a new meaning

  if (existingRow) {
    const oldDef = existingRow[3];
    const oldContext = existingRow[5];

    console.log(`🔍 Duplicate word found. Analyzing context vs existing definition...`);
    
    // Gemini judge is called
    const contextAnalysis = askGeminiIfNewMeaning(wordData, oldDef, oldContext);

    if (!contextAnalysis.is_different) {
      // CASE 1: It's the same -> Reject
      const msg = `The word "${wordData.palabra}" already exists with the same meanign: "${oldDef}". Duplicate won't be created.`;
      console.warn(`⛔ REJECTED: ${msg}`);
      return; // The process stops here
    } else {
      // CASE 2: It's different -> Proceed
      console.log(`✨ New meaning detected (Reason: ${contextAnalysis.reason}). Proceeding as Polysemy.`);
      isPolysemy = true;
    }
  }

  // === 2. BRAIN (Gemini) ===
  let enriched;
  try {
    enriched = callGeminiAnalyst(wordData, isPolysemy);
    console.log("✅ Gemini: Data ready.");
  } catch (err) { console.error("❌ Error Gemini:", err); return; }

  // === 3. AUDIO (Google Cloud TTS) ===
  try {
    const uniqueSuffix = isPolysemy ? "_v2" : ""; //This avoid overwriting the audio in case it already exists
    const wordFilename = `word_${cleanFilename(wordData.palabra)}.mp3`;
    
    // IPA is sent as 4th argument 
    // If Gemini didn't generate the IPA, we send null to use default pronunciation
    enriched.audioWord = callGoogleTTS(wordData.palabra, wordFilename, "word", enriched.ipa); 
    
    if (enriched.ejemplo_raw && wordData.modo !== 'Solo Pronunciación') {
      console.log("🔹 Generating phrase audio with Google tts...");
      const sentenceFilename = `sent_${cleanFilename(wordData.palabra)}.mp3`;
      // Phrases don't carry forced IPA, would be too complex
      enriched.audioSentence = callGoogleTTS(enriched.ejemplo_raw, sentenceFilename, "sentence"); 
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
      console.log(`🎨 Generating visual prompt for: "${wordData.palabra}"...`);
      // GPT-4o-mini refines the Gemini query for DALL-E
      const visualPrompt = generateVisualPrompt(enriched);
      const imgFilename = `img_${cleanFilename(wordData.palabra)}.png`;
      enriched.image = callOpenAIDalle(visualPrompt, imgFilename);
    } else {
      enriched.image = "";
    }
  } catch (err) {
    console.error("⚠️ Error Image:", err.toString());
    enriched.image = ""; 
  }

  // === 5. SAVE ===
  try {
    addToAnkiSheet(enriched);
    console.log("🎉 COMPLETE SUCCESS: Card saved.");
  } catch (err) { console.error("❌ Error Sheets:", err); }
}
// === 2.5 THE JUDGE ===
function askGeminiIfNewMeaning(newData, oldDef, oldCtx) {
  const modelVersion = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;
  
  const prompt = `
    You are a strict linguistic judge avoiding duplicates in a database.
    
    EXISTING ENTRY:
    - Word: "${newData.palabra}"
    - Definition: "${oldDef}"
    - Context used: "${oldCtx}"

    NEW INPUT:
    - Word: "${newData.palabra}"
    - New Context: "${newData.contexto}"

    TASK: 
    Analyze if the New Context implies a SIGNIFICANTLY DIFFERENT meaning (e.g., Phrasal Verb variation, Polysemy, Noun vs Verb) compared to the Existing Definition.
    
    OUTPUT JSON ONLY:
    {
      "is_different": boolean, 
      "reason": "Short explanation (e.g. 'Old is literal, New is idiomatic')"
    }
  `;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const json = JSON.parse(response.getContentText());
    const content = json.candidates[0].content.parts[0].text;
    const cleanJson = content.replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(cleanJson);
  } catch (e) {
    console.error("Judge Error:", e);
    // If the judge fails, we assume it's different, to prevent a data leakage  due to technical error
    return { is_different: true, reason: "Error in check" }; 
  }
}

// === 6. GEMINI ANALYST  ===
function callGeminiAnalyst(wordData, isPolysemy = false) {
  const modelVersion = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${CONFIG.API_KEY}`;

  let promptText = "";
  
  // 🎤 Pronunciation Mode:  cloze is forced
  if (wordData.modo === 'Solo Pronunciación') {
    promptText = `
      You are a linguistic engine. Analyze: "${wordData.palabra}". Context: "${wordData.contexto}".
      TASK: Create Pronunciation data for an Anki Cloze card.
      CRITICAL: You MUST use the cloze format {{c1::word}} in the 'example' field.
      JSON Schema:
      {
        "definition": "IPA transcription (e.g. /wɜːrd/).",
        "ipa": "ˌaɪ.dəmˈpoʊ.tənt",
        "example": "Pronunciation tip for {{c1::${wordData.palabra}}}: [Insert tip regarding stress/linking]",
        "example_raw": "Pronunciation tip for ${wordData.palabra}: [Insert tip regarding stress/linking]",
        "type": "Part of speech",
        "image_query": null
      }
    `;
    // 📚 General vocab mode
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
        "ipa": "ˌaɪ.dəmˈpoʊ.tənt",
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
  // Checks if Gemini forgot the "{{c1::}}"  and fixes it automatically
  rawText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
  const result = JSON.parse(rawText);

  // 🛡️ CLOZE AUTOCORRECTION (UNIVERSAL)
  // if the word is in the phrase, it's enveloped
  let finalExample = result.example;
  // if it's not, we force the format at the beginning so that anki doesn't fail.
  if (finalExample && !finalExample.includes('{{c1::')) {
      const escapedWord = wordData.palabra.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp(`(${escapedWord})`, 'gi');
      if (regex.test(finalExample)) {
          finalExample = finalExample.replace(regex, '{{c1::$1}}');
      } else {
          finalExample = `Note on {{c1::${wordData.palabra}}}: ${finalExample}`;
      }
  }

  let tags = result.tags || "";
  if (isPolysemy) tags = tags ? tags + " polysemy" : "polysemy";

  return {
    ...wordData, 
    definicion: result.definition,
    ipa: result.ipa,
    ejemplo: finalExample, 
    ejemplo_raw: result.example_raw, 
    tipo: result.type,
    tags: tags,
    image_query: result.image_query,
    tag_mode: wordData.modo === 'Solo Pronunciación' ? 'pronunciation' : 'general_vocab'
  };
}

// === 7. OPENAI (DIRECTOR + DALL-E) ===

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

// === 8. TTS & UTILS (Google TTS) ===
function callGoogleTTS(text, filename, type, ipa = null) {
  if (!text) return "";
  
  const url = `https://texttospeech.googleapis.com/v1/text:synthesize?key=${CONFIG.API_KEY}`;
  
  // Configuración de Voz
  let voiceName;
  let useSSML = false;
  let inputPayload = {};

  if (type === "word") {
    voiceName = "en-US-Studio-O"; // Studio soporta SSML perfectamente
    
    // PRONUNCIATION LOGIC INYECTION
    if (ipa) {
      // IPA is cleaned from bars (/word/) in case Gemini put them
      const cleanIPA = ipa.replace(/\//g, '').trim();
      
      // Building SSML
      // <phoneme> forces the model to read the phonetic symbols
      const ssmlText = `<speak><phoneme alphabet="ipa" ph="${cleanIPA}">${text}</phoneme></speak>`;
      
      inputPayload = { ssml: ssmlText };
      useSSML = true;
      console.log(`🔤 Usando IPA forzado para ${text}: [${cleanIPA}]`);
    } else {
      // If no IPA is available, fall back to plain text with manual correction.
      let processedText = text.toLowerCase();
      // Small emergency dictionary in case Gemini breaks/hallucinates
      if (processedText === 'idempotent') processedText = 'eye-dem-po-tent'; 
      inputPayload = { text: processedText };
    }

  } else {
    // For sentences Chirp (Leda) is preferable used, since it's plain text
    voiceName = "en-US-Chirp3-HD-Leda";
    // Period is added for ensuring a good prosody
    let processedSent = text.trim();
    if (!processedSent.endsWith('.')) processedSent += '.';
    inputPayload = { text: processedSent };
  }

  const payload = {
    input: inputPayload,
    voice: { 
      languageCode: "en-US", 
      name: voiceName 
    },
    audioConfig: { 
      audioEncoding: "MP3",
      // Studio (word) allows settings, Chirp (sentence) ignores them
      speakingRate: (type === "word") ? 1.0 : 0.95, 
      pitch: 0.0,
      volumeGainDb: 1.0
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      console.error(`❌ Error API (${voiceName}): ${json.error ? json.error.message : "Desconocido"}`);
      // Fallback: Si falla el SSML (IPA inválido), reintentamos con texto plano
      if (useSSML) {
        console.warn("🔄 Reintentando sin SSML...");
        return callGoogleTTS(text, filename, type, null);
      }
      return "";
    }

    const audioBlob = Utilities.newBlob(Utilities.base64Decode(json.audioContent), "audio/mp3", filename);
    saveFileToDrive(audioBlob, filename, CONFIG.AUDIO_FOLDER_ID);
    return filename;
  } catch (e) {
    console.error(`⚠️ Excepción en TTS: ${e.message}`);
    return "";
  }
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
  const finalTags = data.tags ? data.tags : data.tag_mode;
  sheet.appendRow([
    Utilities.getUuid().substring(0, 8), new Date().toLocaleDateString(), data.palabra, data.definicion, data.ejemplo, data.contexto, data.tipo, 'NO', finalTags, data.audioWord, data.image, data.audioSentence 
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
  SpreadsheetApp.getUi().alert(`✅ The export file is ready :D`);
}
