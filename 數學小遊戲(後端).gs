// --- 設定區 ---
// Sheet IDs — 可在 Script Properties 中覆寫，避免硬編碼
function getConfigId_(propKey, defaultValue) {
  var val = PropertiesService.getScriptProperties().getProperty(propKey);
  return val || defaultValue;
}
function getQuestionsSheetId_() { return getConfigId_('QUESTIONS_SHEET_ID', '1RpKFnPmPDvDvW__nwzZWBnB5jlO6T_aHR9LellTNNiM'); }
function getPrimaryQuestionsSheetId_() { return getConfigId_('PRIMARY_QUESTIONS_SHEET_ID', '1dZaOR-VJi5YcWi36IkXUnr8z958L9bmfTsMd_6Q1vSE'); }
function getNotesSheetId_() { return getConfigId_('NOTES_SHEET_ID', '1i-4_v1u9Q7yhTXESEP4oDdeHCuDEsEpicb3RrUd0oks'); }
function getPrimaryNotesSheetId_() { return getConfigId_('PRIMARY_NOTES_SHEET_ID', '16M6pRAMU1DuBMt0qBltOc1RtMVf7GvList8PRxyc8SY'); }
function getNotesQuizSheetId_() { return getConfigId_('NOTES_QUIZ_SHEET_ID', '1hDpql-EG8zne7NDaabjfuy6Ip1nAxmP-H5I2woEX-tw'); }
function getPrimaryNotesQuizSheetId_() { return getConfigId_('PRIMARY_NOTES_QUIZ_SHEET_ID', '1J6sFDhRLra_A2T7XwPizNRYDWmZ1_HRfy--vkBlk3Qo'); }
function getFirebaseDbUrl_() { return getConfigId_('FIREBASE_DB_URL', 'https://math-game-3747d-default-rtdb.firebaseio.com'); }
function getResultsSpreadsheetId_() { return getConfigId_('RESULTS_SPREADSHEET_ID', '1Ws1Q3uMnpr4erM3vBNbzL7j1a035rNb3pEAUCAqEmO0'); }
const OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions";
const OPENROUTER_MODEL = "google/gemini-2.5-flash";
const OPENROUTER_FALLBACK_MODEL = "meta-llama/llama-3.3-70b-instruct";
const PRIMARY_GRADE_SET_ = { '小一': true, '小二': true, '小三': true, '小四': true, '小五': true, '小六': true };
const SECONDARY_GRADE_SET_ = { '中一': true, '中二': true, '中三': true, '中四': true, '中五': true, '中六': true };

function isPrimaryGrade_(grade) {
  return !!PRIMARY_GRADE_SET_[String(grade || '').trim()];
}

function isSecondaryGrade_(grade) {
  return !!SECONDARY_GRADE_SET_[String(grade || '').trim()];
}

function resolveSchoolLevelFromGrade_(grade) {
  if (isPrimaryGrade_(grade)) return 'primary';
  if (isSecondaryGrade_(grade)) return 'secondary';
  return '';
}

function normalizeSchoolLevel_(schoolLevel, grade) {
  var raw = String(schoolLevel || '').trim().toLowerCase();
  if (raw === '小學') raw = 'primary';
  if (raw === '中學') raw = 'secondary';
  if (raw === 'primary' || raw === 'secondary') return raw;
  return resolveSchoolLevelFromGrade_(grade);
}

function getQuestionsSheetIdForGrade_(grade) {
  return normalizeSchoolLevel_('', grade) === 'primary' ? getPrimaryQuestionsSheetId_() : getQuestionsSheetId_();
}

function getNotesSheetIdForContext_(grade, schoolLevel) {
  return normalizeSchoolLevel_(schoolLevel, grade) === 'primary' ? getPrimaryNotesSheetId_() : getNotesSheetId_();
}

function getNotesQuizSheetIdForContext_(grade, schoolLevel) {
  return normalizeSchoolLevel_(schoolLevel, grade) === 'primary' ? getPrimaryNotesQuizSheetId_() : getNotesQuizSheetId_();
}

function openFirstSheetById_(sheetId, label) {
  if (!sheetId) {
    throw new Error('缺少 ' + label + ' 設定，請檢查 Script Properties。');
  }
  var ss = SpreadsheetApp.openById(sheetId);
  var sheets = ss.getSheets();
  if (!sheets || !sheets.length) {
    throw new Error(label + ' 內沒有可讀取的工作表。');
  }
  return sheets[0];
}

function openQuestionTypeSheetsById_(sheetId, label) {
  if (!sheetId) {
    throw new Error('缺少 ' + label + ' 設定，請檢查 Script Properties。');
  }

  var ss = SpreadsheetApp.openById(sheetId);
  var sheetConfigs = [
    { name: 'multipleChoice', questionType: 'multipleChoice' },
    { name: 'fillBlank', questionType: 'fillBlank' },
    { name: 'matching', questionType: 'matching' }
  ];
  var sources = sheetConfigs.map(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) return null;
    return { sheet: sheet, questionType: config.questionType };
  }).filter(function(item) { return !!item; });

  if (sources.length) return sources;

  var fallbackSheets = ss.getSheets();
  if (!fallbackSheets || !fallbackSheets.length) {
    throw new Error(label + ' 內沒有可讀取的工作表。');
  }

  return [{ sheet: fallbackSheets[0], questionType: '' }];
}

function doGet() {
  return HtmlService.createTemplateFromFile('數學小遊戲')
    .evaluate()
    .setTitle("數學安多fun - 終極動態版")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// Firebase REST API 寫入輔助函式
// ==========================================
function pushToFirebase_(path, payload) {
  try {
    var cleanPath = String(path || '').replace(/^\/+|\/+$/g, '');
    var url = getFirebaseDbUrl_() + '/' + cleanPath + '.json';
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload || {}),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      return { error: 'Firebase 寫入失敗，HTTP ' + code + '：' + res.getContentText() };
    }
    var body = {};
    try { body = JSON.parse(res.getContentText()); } catch (e) {}
    return { success: true, name: body && body.name ? body.name : null };
  } catch (e) {
    return { error: 'Firebase 寫入例外: ' + e.message };
  }
}

function fetchFromFirebase_(path) {
  try {
    var cleanPath = String(path || '').replace(/^\/+|\/+$/g, '');
    var url = getFirebaseDbUrl_() + '/' + cleanPath + '.json';
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      return { error: 'Firebase 讀取失敗，HTTP ' + code + '：' + res.getContentText() };
    }
    try {
      return { success: true, data: JSON.parse(res.getContentText() || 'null') };
    } catch (e) {
      return { error: 'Firebase 回傳格式錯誤：' + e.message };
    }
  } catch (e) {
    return { error: 'Firebase 讀取例外: ' + e.message };
  }
}

function patchFirebase_(path, payload) {
  try {
    var cleanPath = String(path || '').replace(/^\/+|\/+$/g, '');
    var url = getFirebaseDbUrl_() + '/' + cleanPath + '.json';
    var res = UrlFetchApp.fetch(url, {
      method: 'patch',
      contentType: 'application/json',
      payload: JSON.stringify(payload || {}),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      return { error: 'Firebase 更新失敗，HTTP ' + code + '：' + res.getContentText() };
    }
    return { success: true };
  } catch (e) {
    return { error: 'Firebase 更新例外: ' + e.message };
  }
}

// ==========================================
// 通用欄位搜尋輔助
// ==========================================
function normalizeHeaders_(rawHeaders) {
  return rawHeaders.map(function(h) {
    return String(h || "").toLowerCase().normalize('NFKC').replace(/[\s_\.\-\:：]+/g, '');
  });
}

function findColumnIndex_(headers, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var idx = headers.indexOf(possibleNames[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function shuffleArray_(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = arr[i];
    arr[i] = arr[j];
    arr[j] = temp;
  }
  return arr;
}

function normalizeQuestionType_(rawType) {
  var value = String(rawType || '').toLowerCase().normalize('NFKC').replace(/[\s_\.\-:：]+/g, '');
  if (!value) return 'multipleChoice';
  if (['fillblank', 'fillintheblank', 'blank', 'input', 'textinput', '填充題', '填空題', '填充', '填空'].indexOf(value) !== -1) {
    return 'fillBlank';
  }
  if (['matching', 'match', 'pairing', 'pairmatch', '配對題', '配對'].indexOf(value) !== -1) {
    return 'matching';
  }
  return 'multipleChoice';
}

function parseCellList_(value) {
  return String(value || '')
    .split(/\r?\n|[\|｜;；]+/)
    .map(function(item) { return String(item || '').trim(); })
    .filter(function(item) { return item !== ''; });
}

function buildAcceptedAnswers_(answerValue, extraValue) {
  var accepted = [];
  var pushUnique = function(item) {
    var normalized = String(item || '').trim();
    if (!normalized || accepted.indexOf(normalized) !== -1) return;
    accepted.push(normalized);
  };

  pushUnique(answerValue);
  parseCellList_(extraValue).forEach(pushUnique);
  return accepted;
}

function extractQuestionOptions_(row, optIndices) {
  var options = [];
  (optIndices || []).forEach(function(idx) {
    if (idx < row.length && String(row[idx]).trim() !== '') {
      options.push(String(row[idx]).trim());
    }
  });

  if (options.length === 1) {
    var singleOpt = options[0];
    if (singleOpt.includes(',') || singleOpt.includes('，')) {
      options = singleOpt.split(/[,，]/).map(function(item) { return item.trim(); }).filter(function(item) { return item !== ''; });
    } else if (singleOpt.includes('\n')) {
      options = singleOpt.split('\n').map(function(item) { return item.trim(); }).filter(function(item) { return item !== ''; });
    } else if (/A[\.\)]/i.test(singleOpt) && /B[\.\)]/i.test(singleOpt)) {
      var temp = singleOpt
        .replace(/A[\.\)]/gi, '||A.')
        .replace(/B[\.\)]/gi, '||B.')
        .replace(/C[\.\)]/gi, '||C.')
        .replace(/D[\.\)]/gi, '||D.');
      options = temp.split('||').map(function(item) { return item.trim(); }).filter(function(item) { return item !== ''; });
    } else {
      var spaceSplit = singleOpt.split(/\s+/).map(function(item) { return item.trim(); }).filter(function(item) { return item !== ''; });
      if (spaceSplit.length > 1) options = spaceSplit;
    }
  }

  return options;
}

function buildMatchingPairs_(leftValue, rightValue) {
  var leftItems = parseCellList_(leftValue);
  var rightItems = parseCellList_(rightValue);
  var count = Math.min(leftItems.length, rightItems.length);
  var pairs = [];

  for (var i = 0; i < count; i++) {
    pairs.push({ left: leftItems[i], right: rightItems[i] });
  }

  return pairs;
}

function buildQuestionPayload_(row, config) {
  var questionType = config.forcedQuestionType || normalizeQuestionType_(config.typeIdx !== -1 ? row[config.typeIdx] : '');
  var answer = String(row[config.ansIdx] || '').trim();
  var acceptedAnswers = buildAcceptedAnswers_(answer, config.acceptedIdx !== -1 ? row[config.acceptedIdx] : '');
  var payload = {
    question: String(row[config.qIdx] || '').trim(),
    questionType: questionType,
    answer: answer,
    acceptedAnswers: acceptedAnswers,
    explanation: (config.expIdx !== -1 && String(row[config.expIdx]).trim() !== '') ? String(row[config.expIdx]).trim() : (config.defaultExplanation || ''),
    imageUrl: (config.imgIdx !== -1 && String(row[config.imgIdx] || '').trim() !== '') ? String(row[config.imgIdx]).trim() : ''
  };

  if (typeof config.pointsValue !== 'undefined') {
    payload.points = config.pointsValue;
  }

  if (questionType === 'matching') {
    payload.matchingPairs = buildMatchingPairs_(
      config.matchLeftIdx !== -1 ? row[config.matchLeftIdx] : '',
      config.matchRightIdx !== -1 ? row[config.matchRightIdx] : ''
    );
    payload.options = [];
    if (!payload.matchingPairs.length) return null;
    return payload;
  }

  if (questionType === 'fillBlank') {
    payload.options = [];
    return payload;
  }

  payload.options = extractQuestionOptions_(row, config.optIndices);
  if (payload.options.length === 0) payload.options = ['選項讀取失敗'];
  shuffleArray_(payload.options);
  return payload;
}

function selectQuestionsByComposition_(questions, limit, desiredSpecialCounts) {
  var maxQuestions = Math.max(0, Number(limit) || 10);
  var buckets = {
    multipleChoice: [],
    fillBlank: [],
    matching: []
  };
  var selected = [];
  var specialCounts = desiredSpecialCounts || {};

  (questions || []).forEach(function(question) {
    var questionType = normalizeQuestionType_(question && question.questionType);
    if (!buckets[questionType]) buckets[questionType] = [];
    buckets[questionType].push(question);
  });

  Object.keys(buckets).forEach(function(questionType) {
    shuffleArray_(buckets[questionType]);
  });

  Object.keys(specialCounts).forEach(function(questionType) {
    var bucket = buckets[questionType] || [];
    var takeCount = Math.min(specialCounts[questionType], bucket.length, Math.max(0, maxQuestions - selected.length));
    if (takeCount > 0) {
      selected = selected.concat(bucket.splice(0, takeCount));
    }
  });

  var multipleChoiceBucket = buckets.multipleChoice || [];
  var remainingSlots = Math.max(0, maxQuestions - selected.length);
  if (remainingSlots > 0 && multipleChoiceBucket.length) {
    selected = selected.concat(multipleChoiceBucket.splice(0, remainingSlots));
  }

  remainingSlots = Math.max(0, maxQuestions - selected.length);
  if (remainingSlots > 0) {
    var fallbackPool = [];
    ['fillBlank', 'matching', 'multipleChoice'].forEach(function(questionType) {
      if (buckets[questionType] && buckets[questionType].length) {
        fallbackPool = fallbackPool.concat(buckets[questionType]);
      }
    });
    if (fallbackPool.length) {
      shuffleArray_(fallbackPool);
      selected = selected.concat(fallbackPool.slice(0, remainingSlots));
    }
  }

  shuffleArray_(selected);
  return selected.slice(0, maxQuestions);
}

function selectChallengeQuestions_(questions, limit) {
  return selectQuestionsByComposition_(questions, limit, {
    fillBlank: 3,
    matching: 2
  });
}

function selectNotesQuizQuestions_(questions, limit) {
  return selectQuestionsByComposition_(questions, limit, {
    fillBlank: 1,
    matching: 1
  });
}

function formatResultTime_() {
  var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
}

function buildStudentResultPlayerKey_(playerInfo) {
  playerInfo = playerInfo || {};
  var className = String(playerInfo.class || '').trim();
  var number = String(playerInfo.number || '').trim();
  return className && number ? (className + '-' + number) : '';
}

function buildStudentResultIndexKey_(mode, resultId) {
  return String(mode || 'result').trim() + '__' + String(resultId || '').trim();
}

function getStudentChallengeResultsPath_(playerKey) {
  return 'students/' + String(playerKey || '').trim() + '/challengeResults';
}

function extractStudentResultIndexEntries_(resultRoot, playerInfo, mode) {
  var entries = {};
  var data = resultRoot && typeof resultRoot === 'object' ? resultRoot : {};
  var targetKey = buildStudentResultPlayerKey_(playerInfo);

  Object.keys(data).forEach(function(resultId) {
    var item = data[resultId] || {};
    if (buildStudentResultPlayerKey_(item.player) !== targetKey) return;
    entries[buildStudentResultIndexKey_(mode, resultId)] = item;
  });

  return entries;
}

function writeStudentResultIndex_(playerInfo, mode, resultId, payload) {
  var playerKey = buildStudentResultPlayerKey_(playerInfo || (payload && payload.player) || {});
  if (!playerKey || !resultId) return { error: '缺少學生資料或 resultId，未能建立結果索引。' };

  var indexPayload = {};
  indexPayload[buildStudentResultIndexKey_(mode, resultId)] = payload || {};

  var indexRes = patchFirebase_(getStudentChallengeResultsPath_(playerKey), indexPayload);
  if (indexRes.error) return indexRes;

  var metaRes = patchFirebase_('students/' + playerKey + '/resultIndexMeta', {
    updatedAt: Date.now()
  });
  if (metaRes.error) Logger.log('[writeStudentResultIndex_] ' + metaRes.error);

  var profileInvalidateRes = patchFirebase_('students/' + playerKey + '/profile', {
    generatedAt: 0
  });
  if (profileInvalidateRes.error) Logger.log('[writeStudentResultIndex_] ' + profileInvalidateRes.error);

  return { success: true };
}

function migrateStudentResultsIndexFromLegacy_(playerInfo) {
  var playerKey = buildStudentResultPlayerKey_(playerInfo);
  if (!playerKey) return { error: '缺少學生資料，未能回填結果索引。' };

  var singleRes = fetchFromFirebase_('results/single');
  if (singleRes.error) return { error: singleRes.error };
  var multiRes = fetchFromFirebase_('results/multi');
  if (multiRes.error) return { error: multiRes.error };

  var indexedEntries = {};
  var singleEntries = extractStudentResultIndexEntries_(singleRes.data, playerInfo, 'single');
  var multiEntries = extractStudentResultIndexEntries_(multiRes.data, playerInfo, 'multi');

  Object.keys(singleEntries).forEach(function(key) { indexedEntries[key] = singleEntries[key]; });
  Object.keys(multiEntries).forEach(function(key) { indexedEntries[key] = multiEntries[key]; });

  if (Object.keys(indexedEntries).length) {
    var indexRes = patchFirebase_(getStudentChallengeResultsPath_(playerKey), indexedEntries);
    if (indexRes.error) return { error: indexRes.error };
  }

  var metaRes = patchFirebase_('students/' + playerKey + '/resultIndexMeta', {
    version: 1,
    migratedAt: Date.now(),
    updatedAt: Date.now()
  });
  if (metaRes.error) Logger.log('[migrateStudentResultsIndexFromLegacy_] ' + metaRes.error);

  return { success: true, data: indexedEntries };
}

function saveToSheet(data, type) {
  try {
    var t = String(type || '').toLowerCase();
    var sheetName = t === 'multi' ? 'results(multiplayers)' : 'results(single player)';
    var ss = SpreadsheetApp.openById(getResultsSpreadsheetId_());
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: '找不到工作表：' + sheetName };

    var d = data || {};
    var p = d.player || {};
    var topic = String(d.topic || d.theme || d.subject || '');
    var difficulty = String(d.difficulty || '');
    var accuracy = Number(d.accuracyPct || d.accuracy || d.correctRate || 0);
    var rank = Number(d.rank || d.ranking || 0);

    if (t === 'multi') {
      // 完成時間, 班別, 學號, 課題, 難度, 分數, 答對百分比, 排名
      sheet.appendRow([
        formatResultTime_(),
        String(p.class || ''),
        String(p.number || ''),
        topic,
        difficulty,
        Number(d.score || 0),
        accuracy,
        rank
      ]);
    } else {
      // 完成時間, 班別, 學號, 課題, 難度, 分數, 答對百分比
      sheet.appendRow([
        formatResultTime_(),
        String(p.class || ''),
        String(p.number || ''),
        topic,
        difficulty,
        Number(d.score || 0),
        accuracy
      ]);
    }
    return { success: true };
  } catch (e) {
    return { error: 'Google Sheet 寫入失敗: ' + e.message };
  }
}

// ==========================================
// 1. 取得不重複的遊戲主題
// ==========================================
function getAvailableThemes(grade) {
  try {
    const sheet = openFirstSheetById_(getQuestionsSheetIdForGrade_(grade), '題庫 Sheet ID');
    const data = sheet.getDataRange().getValues();

    const headers = normalizeHeaders_(data.shift());
    const themeIdx = findColumnIndex_(headers, ['theme', '主題', '類型', 'category', 'topic']);
    const gradeIdx = findColumnIndex_(headers, ['級別', 'grade', '年級', 'level']);

    if (themeIdx === -1) return [];
    const themes = [];
    for (let i = 0; i < data.length; i++) {
      if (grade && gradeIdx !== -1) {
        const rowGrade = String(data[i][gradeIdx] || "").trim();
        if (rowGrade !== grade) continue;
      }
      const theme = String(data[i][themeIdx] || "").trim();
      if (theme !== "" && themes.indexOf(theme) === -1) {
        themes.push(theme);
      }
    }
    return themes;
  } catch (e) {
    Logger.log("讀取主題錯誤: " + e.message);
    return [];
  }
}

// ==========================================
// 2. 讀取遊戲題庫 (Google Sheets)
// ==========================================
function getQuestions(difficulty, theme, grade) {
  try {
    const difficultyMap = { '簡單': 'easy', '普通': 'medium', '中等': 'medium', '困難': 'hard' };
    const targetDiff = difficultyMap[difficulty] || difficulty;
    const sources = openQuestionTypeSheetsById_(getQuestionsSheetIdForGrade_(grade), '題庫 Sheet ID');
    let questions = [];

    sources.forEach(function(source) {
      const data = source.sheet.getDataRange().getValues();
      if (!data || data.length < 2) return;

      const headers = normalizeHeaders_(data.shift());
      const findIdx = function(names) { return findColumnIndex_(headers, names); };
      const themeIdx = findIdx(['theme', '主題', '類型', 'category', 'topic']);
      const diffIdx  = findIdx(['difficulty', '難度', 'level']);
      const gradeIdx = findIdx(['級別', 'grade', '年級']);
      const qIdx     = findIdx(['question', '題目', '問題', 'q']);
      const ansIdx   = findIdx(['answer', '答案', '解答', 'correctanswer', 'ans']);
      const expIdx   = findIdx(['explanation', '詳解', '解釋', '說明', '解析', 'exp']);
      const typeIdx  = findIdx(['questiontype', '題型', '題目類型', 'qtype', 'questionkind']);
      const acceptedIdx = findIdx(['acceptedanswers', 'acceptableanswers', 'altanswers', '可接受答案', '其他答案']);
      const imgIdx   = findIdx(['image', 'imageurl', '圖片', '圖片url', '插圖']);
      const matchLeftIdx = findIdx(['matchingleft', 'leftitems', 'leftoptions', 'pairleft', '配對左', '配對左欄', '左欄']);
      const matchRightIdx = findIdx(['matchingright', 'rightitems', 'rightoptions', 'pairright', '配對右', '配對右欄', '右欄']);
      const optAIdx  = findIdx(['optiona', '選項a', 'a', 'option1', '選項1', 'choicea', 'choice1', '選擇a', '選擇1', 'opt1', 'opta', '1', '選項', '選擇', 'option', 'options', 'choices']);
      const optBIdx  = findIdx(['optionb', '選項b', 'b', 'option2', '選項2', 'choiceb', 'choice2', '選擇b', '選擇2', 'opt2', 'optb', '2']);
      const optCIdx  = findIdx(['optionc', '選項c', 'c', 'option3', '選項3', 'choicec', 'choice3', '選擇c', '選擇3', 'opt3', 'optc', '3']);
      const optDIdx  = findIdx(['optiond', '選項d', 'd', 'option4', '選項4', 'choiced', 'choice4', '選擇d', '選擇4', 'opt4', 'optd', '4']);

      if (qIdx === -1 || ansIdx === -1) return;

      let optIndices = [];
      if (optAIdx !== -1) optIndices.push(optAIdx);
      if (optBIdx !== -1 && !optIndices.includes(optBIdx)) optIndices.push(optBIdx);
      if (optCIdx !== -1 && !optIndices.includes(optCIdx)) optIndices.push(optCIdx);
      if (optDIdx !== -1 && !optIndices.includes(optDIdx)) optIndices.push(optDIdx);

      if (optIndices.length === 1) {
        let firstOpt = optIndices[0];
        for (let i = 1; i <= 3; i++) {
          let nextIdx = firstOpt + i;
          if (nextIdx < headers.length && (headers[nextIdx] === "" || headers[nextIdx] === headers[firstOpt])) {
            if (!optIndices.includes(nextIdx)) optIndices.push(nextIdx);
          }
        }
      }

      if (optIndices.length === 0 && qIdx !== -1 && ansIdx !== -1) {
        let start = Math.min(qIdx, ansIdx) + 1;
        let end = Math.max(qIdx, ansIdx);
        for (let i = start; i < end; i++) optIndices.push(i);
      }

      data.forEach(function(row) {
        const rowDiff = diffIdx !== -1 ? String(row[diffIdx] || "").trim().toLowerCase() : "";
        const rowTheme = themeIdx !== -1 ? String(row[themeIdx] || "").trim() : "";
        const rowGrade = gradeIdx !== -1 ? String(row[gradeIdx] || "").trim() : "";
        const diffMatch = (!targetDiff || targetDiff === '全部') ? true : (rowDiff === targetDiff);
        const themeMatch = (!theme || theme === '全部') ? true : (rowTheme.toLowerCase() === String(theme).toLowerCase());
        const gradeMatch = (!grade || grade === '全部') ? true : (rowGrade === grade);
        const isNotEmpty = String(row[qIdx]).trim() !== "";
        if (!diffMatch || !themeMatch || !gradeMatch || !isNotEmpty) return;

        const payload = buildQuestionPayload_(row, {
          qIdx: qIdx,
          ansIdx: ansIdx,
          expIdx: expIdx,
          imgIdx: imgIdx,
          typeIdx: typeIdx,
          acceptedIdx: acceptedIdx,
          matchLeftIdx: matchLeftIdx,
          matchRightIdx: matchRightIdx,
          optIndices: optIndices,
          forcedQuestionType: source.questionType,
          defaultExplanation: '請留意計算步驟喔！'
        });
        if (payload) questions.push(payload);
      });
    });

    if (questions.length === 0) {
      return { error: '找不到符合條件的題目 (難度: ' + difficulty + ', 主題: ' + theme + ')' };
    }

    questions = selectChallengeQuestions_(questions, 12);

    if (questions.length === 0) {
      return { error: '符合條件的題目缺少必要欄位，請檢查 questionType 與題型資料。' };
    }

    return questions;
  } catch (e) {
    return { error: '系統錯誤: ' + e.message };
  }
}

// ==========================================
// 分數合理性驗證
// ==========================================
function validateScore_(score, meta) {
  var maxQuestions = 12;
  var basePoints = 10;
  var maxTimeBonus = 60; // timeLeft/2 where max timeLeft = 120
  var maxComboBonus = 5 * maxQuestions; // 5 * combo at max
  var maxPerQuestion = (basePoints + maxTimeBonus + maxComboBonus) * 1.5; // with COMBO_SCORE_MULTIPLIER
  var maxPossible = maxPerQuestion * maxQuestions;
  var s = Number(score || 0);
  if (s < 0 || s > maxPossible || !isFinite(s)) return false;
  return true;
}

// ==========================================
// 3. 儲存單人模式成績 (Firebase)
// ==========================================
function saveSingleResult(playerInfo, score, meta) {
  try {
    if (!validateScore_(score, meta)) {
      return { error: '分數超出合理範圍，無法儲存。' };
    }
    meta = meta || {};
    var payload = {
      mode: 'single',
      createdAt: new Date().toISOString(),
      player: {
        class: playerInfo.class || '',
        number: playerInfo.number || '',
        avatar: playerInfo.avatar || '',
        avatarName: playerInfo.avatarName || ''
      },
      topic: String(meta.topic || meta.theme || ''),
      difficulty: String(meta.difficulty || ''),
      accuracyPct: Number(meta.accuracyPct || meta.accuracy || 0),
      score: Number(score || 0)
    };
    var fbRes = pushToFirebase_('results/single', payload);
    if (fbRes && fbRes.success && fbRes.name) {
      var indexRes = writeStudentResultIndex_(playerInfo, 'single', fbRes.name, payload);
      if (indexRes.error) Logger.log('[saveSingleResult] 學生結果索引寫入失敗（已忽略，不影響主紀錄）: ' + indexRes.error);
    }
    var sheetRes = saveToSheet(payload, 'single');
    if (sheetRes && sheetRes.error) {
      Logger.log('[saveSingleResult] Sheet 寫入失敗（已忽略，不影響 Firebase）: ' + sheetRes.error);
    }
    return fbRes;
  } catch (e) {
    return { error: e.message };
  }
}

// ==========================================
// 4. 多人模式結算成績 (Firebase)
// ==========================================
function finishGame(roomId, playerInfo, score, meta) {
  try {
    if (!validateScore_(score, meta)) {
      return { error: '分數超出合理範圍，無法儲存。' };
    }
    meta = meta || {};
    var payload = {
      mode: 'multi',
      roomId: String(roomId || ''),
      createdAt: new Date().toISOString(),
      player: {
        class: playerInfo && playerInfo.class ? playerInfo.class : '',
        number: playerInfo && playerInfo.number ? playerInfo.number : '',
        avatar: playerInfo && playerInfo.avatar ? playerInfo.avatar : '',
        avatarName: playerInfo && playerInfo.avatarName ? playerInfo.avatarName : ''
      },
      topic: String(meta.topic || meta.theme || ''),
      difficulty: String(meta.difficulty || ''),
      accuracyPct: Number(meta.accuracyPct || meta.accuracy || 0),
      rank: Number(meta.rank || meta.ranking || 0),
      score: Number(score || 0)
    };
    var fbRes = pushToFirebase_('results/multi', payload);
    if (fbRes && fbRes.success && fbRes.name) {
      var indexRes = writeStudentResultIndex_(playerInfo, 'multi', fbRes.name, payload);
      if (indexRes.error) Logger.log('[finishGame] 學生結果索引寫入失敗（已忽略，不影響主紀錄）: ' + indexRes.error);
    }
    var sheetRes = saveToSheet(payload, 'multi');
    if (sheetRes && sheetRes.error) {
      Logger.log('[finishGame] Sheet 寫入失敗（已忽略，不影響 Firebase）: ' + sheetRes.error);
    }
    return fbRes;
  } catch (e) {
    return { error: e.message };
  }
}

function inferGradeFromClass_(className) {
  var raw = String(className || '').trim();
  if (!raw) return '';
  if (/^小[一二三四五六]/.test(raw)) return raw.slice(0, 2);
  if (/^中[一二三四五六]/.test(raw)) return raw.slice(0, 2);
  var match = raw.match(/^(\d+)/);
  if (!match) return '';
  var map = {
    '1': '中一',
    '2': '中二',
    '3': '中三',
    '4': '中四',
    '5': '中五',
    '6': '中六'
  };
  return map[match[1]] || '';
}

function sanitizeFirebaseKey_(value) {
  return String(value || '').replace(/[.#$\[\]\/]/g, '_');
}

function restoreTopicDisplayLabel_(value) {
  return String(value || '').replace(/_/g, ' ').trim();
}

function toTimestamp_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  if (!value) return 0;
  var stamp = Date.parse(value);
  return isFinite(stamp) ? stamp : 0;
}

function clampNumber_(value, min, max) {
  var num = Number(value || 0);
  if (!isFinite(num)) num = 0;
  if (typeof min === 'number' && num < min) num = min;
  if (typeof max === 'number' && num > max) num = max;
  return num;
}

function buildNotesTopicMeta_(grade, schoolLevel) {
  var result = {
    topics: [],
    sectionCountByTopic: {},
    topicLabelBySanitized: {}
  };

  try {
    var sheet = openFirstSheetById_(getNotesSheetIdForContext_(grade, schoolLevel), '溫習筆記 Sheet ID');
    var data = sheet.getDataRange().getValues();
    var headers = normalizeHeaders_(data.shift());
    var topicIdx = findColumnIndex_(headers, ['課題', 'topic', '主題', 'theme', 'chapter']);
    var gradeIdx = findColumnIndex_(headers, ['級別', 'grade', '年級', 'level', '班級']);
    var sectionIdx = findColumnIndex_(headers, ['section', '章節', '單元']);
    if (topicIdx === -1) return result;

    var topicSeen = {};
    var sectionSeenByTopic = {};
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (grade && gradeIdx !== -1 && String(row[gradeIdx] || '').trim() !== grade) continue;
      var topic = String(row[topicIdx] || '').trim();
      if (!topic) continue;
      var sanitized = sanitizeFirebaseKey_(topic);
      result.topicLabelBySanitized[sanitized] = topic;
      if (!topicSeen[topic]) {
        topicSeen[topic] = true;
        result.topics.push(topic);
      }
      if (!sectionSeenByTopic[topic]) sectionSeenByTopic[topic] = {};
      var section = sectionIdx !== -1 ? String(row[sectionIdx] || '').trim() : '';
      if (section) sectionSeenByTopic[topic][section] = true;
    }

    result.topics.forEach(function(topic) {
      var count = Object.keys(sectionSeenByTopic[topic] || {}).length;
      result.sectionCountByTopic[topic] = count;
    });
  } catch (e) {
    Logger.log('[buildNotesTopicMeta_] ' + e.message);
  }

  return result;
}

function normalizeTopicLabel_(rawTopic, topicMeta) {
  var label = String(rawTopic || '').trim();
  if (!label) return '';
  var restored = restoreTopicDisplayLabel_(label);
  if (topicMeta && topicMeta.topicLabelBySanitized && topicMeta.topicLabelBySanitized[label]) {
    return topicMeta.topicLabelBySanitized[label];
  }
  if (topicMeta && topicMeta.topicLabelBySanitized && topicMeta.topicLabelBySanitized[sanitizeFirebaseKey_(restored)]) {
    return topicMeta.topicLabelBySanitized[sanitizeFirebaseKey_(restored)];
  }
  return restored;
}

function buildChallengeEntries_(resultRoot, playerInfo) {
  var entries = [];
  var data = resultRoot && typeof resultRoot === 'object' ? resultRoot : {};
  var targetClass = String(playerInfo && playerInfo.class || '').trim();
  var targetNumber = String(playerInfo && playerInfo.number || '').trim();

  Object.keys(data).forEach(function(key) {
    var item = data[key] || {};
    var player = item.player || {};
    if (String(player.class || '').trim() !== targetClass || String(player.number || '').trim() !== targetNumber) return;
    entries.push({
      topic: String(item.topic || item.theme || ''),
      difficulty: String(item.difficulty || ''),
      accuracyPct: clampNumber_(item.accuracyPct || item.accuracy || 0, 0, 100),
      score: clampNumber_(item.score || 0, 0),
      createdAt: toTimestamp_(item.createdAt || item.timestamp || 0)
    });
  });

  return entries;
}

function buildAchievementCatalog_(totalTopicCount) {
  return [
    { id: 'first_steps', title: '踏出首步', description: '完成第一場數學挑戰', icon: '🎮', target: 1, type: 'challenges' },
    { id: 'math_enthusiast', title: '熱血解題王', description: '完成 10 場數學挑戰', icon: '🔥', target: 10, type: 'challenges' },
    { id: 'perfect_score', title: '百發百中', description: '最佳成績達到 100 分', icon: '💯', target: 100, type: 'bestScore' },
    { id: 'book_worm', title: '知識探索者', description: '完成 3 個數學課題', icon: '📚', target: 3, type: 'topics' },
    { id: 'collector', title: '尋寶獵人', description: '兌換 5 個角色或特效', icon: '🎁', target: 5, type: 'redeemed' },
    { id: 'math_master', title: '登峰造極', description: '完成全部數學課題', icon: '👑', target: Math.max(1, Number(totalTopicCount || 0)), type: 'allTopics' }
  ];
}

function computeAchievementProgress_(achievement, metrics) {
  if (!achievement) return { current: 0, target: 1, unlocked: false };
  var current = 0;
  if (achievement.type === 'challenges') current = Number(metrics.completedChallenges || 0);
  else if (achievement.type === 'bestScore') current = Number(metrics.bestScore || 0);
  else if (achievement.type === 'topics') current = Number(metrics.completedTopics || 0);
  else if (achievement.type === 'redeemed') current = Number(metrics.redeemedRoles || 0);
  else if (achievement.type === 'allTopics') current = Number(metrics.completedTopics || 0);
  var target = Math.max(1, Number(achievement.target || 1));
  var unlocked = current >= target;
  return {
    current: current,
    target: target,
    unlocked: unlocked,
    progressPct: Math.round((Math.min(current, target) / target) * 100)
  };
}

function buildProfileRecommendation_(metrics, masteryTopics, topicMeta) {
  var topics = Array.isArray(masteryTopics) ? masteryTopics.slice() : [];
  var weakest = topics.filter(function(item) { return item.dataPoints > 0; }).sort(function(a, b) {
    return a.score - b.score || a.label.localeCompare(b.label);
  })[0] || null;

  var unfinished = (topicMeta && Array.isArray(topicMeta.topics) ? topicMeta.topics : []).find(function(topic) {
    return metrics.completedTopicLabels.indexOf(topic) === -1;
  }) || '';

  if (weakest && weakest.notesCompletionPct < 100) {
    return {
      title: '先補完整個弱項課題',
      description: '「' + weakest.label + '」仍有未完成的溫習章節，先把課題讀完並完成測驗，掌握度會提升得最快。',
      cta: '📖 去完成課題',
      action: 'notes',
      topic: weakest.label
    };
  }
  if (weakest && weakest.gameAttempts > 0 && weakest.gameAccuracy < 70) {
    return {
      title: '先把挑戰失分課題補強',
      description: '「' + weakest.label + '」在遊戲挑戰中的準確度較低，建議先回顧後再打一場同課題挑戰。',
      cta: '🎮 再挑戰弱項',
      action: 'challenge',
      topic: weakest.label
    };
  }
  if (unfinished) {
    return {
      title: '繼續解鎖下一個課題',
      description: '你還未完成「' + unfinished + '」，先完成這個課題可以同時推進成就與整體掌握度。',
      cta: '📘 前往未完成課題',
      action: 'notes',
      topic: unfinished
    };
  }
  if (Number(metrics.completedChallenges || 0) < 10) {
    return {
      title: '再累積幾場實戰挑戰',
      description: '目前挑戰場數仍有空間，先用幾場短挑戰把不同課題的實戰表現拉齊。',
      cta: '⚔️ 再玩一場',
      action: 'challenge',
      topic: ''
    };
  }
  return {
    title: '狀態穩定，適合衝更高分',
    description: '整體掌握度已開始穩定，下一步適合用更高難度或多人對戰檢驗真正熟練度。',
    cta: '🚀 挑戰更高難度',
    action: 'challenge',
    topic: ''
  };
}

function getStudentProfileData(playerInfo, forceRefresh) {
  try {
    playerInfo = playerInfo || {};
    var className = String(playerInfo.class || '').trim();
    var number = String(playerInfo.number || '').trim();
    if (!className || !number) return { error: '缺少學生資料，未能整理 profile。' };

    var playerKey = className + '-' + number;
    var studentRes = fetchFromFirebase_('students/' + playerKey);
    if (studentRes.error) return { error: studentRes.error };
    var student = studentRes.data || {};
    if (!forceRefresh && student.profile && student.profile.generatedAt && (Date.now() - Number(student.profile.generatedAt || 0) < 90000)) {
      return student.profile;
    }

    var grade = inferGradeFromClass_(className);
    var topicMeta = buildNotesTopicMeta_(grade);
    var notesProgress = student.notesProgress || {};
    var purchased = (student.shop && student.shop.purchased) || {};
    var indexedResultRoot = student.challengeResults || {};
    var resultIndexMeta = student.resultIndexMeta || {};
    if (Number(resultIndexMeta.version || 0) < 1) {
      var migrationRes = migrateStudentResultsIndexFromLegacy_(playerInfo);
      if (migrationRes.error) return { error: migrationRes.error };
      indexedResultRoot = migrationRes.data || indexedResultRoot;
    }

    var entries = buildChallengeEntries_(indexedResultRoot, playerInfo);
    entries.sort(function(a, b) { return b.createdAt - a.createdAt; });

    var topicStats = {};
    entries.forEach(function(entry) {
      var label = normalizeTopicLabel_(entry.topic, topicMeta);
      if (!label) return;
      if (!topicStats[label]) {
        topicStats[label] = {
          label: label,
          gameAttempts: 0,
          gameAccuracyTotal: 0,
          gameBestScore: 0,
          notesAccuracy: 0,
          notesCompletionPct: 0,
          notesCompleted: false,
          passedSections: 0,
          totalSections: Number(topicMeta.sectionCountByTopic[label] || 0)
        };
      }
      topicStats[label].gameAttempts += 1;
      topicStats[label].gameAccuracyTotal += clampNumber_(entry.accuracyPct, 0, 100);
      topicStats[label].gameBestScore = Math.max(topicStats[label].gameBestScore, clampNumber_(entry.score, 0));
    });

    Object.keys(notesProgress).forEach(function(rawKey) {
      var item = notesProgress[rawKey] || {};
      var label = normalizeTopicLabel_(rawKey, topicMeta);
      if (!label) return;
      if (!topicStats[label]) {
        topicStats[label] = {
          label: label,
          gameAttempts: 0,
          gameAccuracyTotal: 0,
          gameBestScore: 0,
          notesAccuracy: 0,
          notesCompletionPct: 0,
          notesCompleted: false,
          passedSections: 0,
          totalSections: Number(topicMeta.sectionCountByTopic[label] || 0)
        };
      }
      var stat = topicStats[label];
      var sections = item.sections || {};
      var passedSections = 0;
      Object.keys(sections).forEach(function(sectionKey) {
        if (sections[sectionKey] && sections[sectionKey].passed) passedSections++;
      });
      stat.passedSections = Math.max(stat.passedSections, passedSections);
      stat.totalSections = Math.max(stat.totalSections, Number(topicMeta.sectionCountByTopic[label] || 0), passedSections);
      var quizTotal = clampNumber_(item.quizTotal || 0, 0);
      var quizScore = clampNumber_(item.quizScore || 0, 0);
      stat.notesAccuracy = quizTotal > 0 ? Math.round((quizScore / quizTotal) * 100) : stat.notesAccuracy;
      var completionPct = item.topicComplete
        ? 100
        : (stat.totalSections > 0 ? Math.round((passedSections / stat.totalSections) * 100) : (passedSections > 0 ? 50 : 0));
      stat.notesCompletionPct = Math.max(stat.notesCompletionPct, completionPct);
      stat.notesCompleted = !!item.topicComplete;
    });

    var masteryTopics = Object.keys(topicStats).map(function(label) {
      var stat = topicStats[label];
      var hasNotes = stat.notesAccuracy > 0 || stat.notesCompletionPct > 0 || stat.notesCompleted;
      var hasGame = stat.gameAttempts > 0;
      var gameAccuracy = stat.gameAttempts > 0 ? Math.round(stat.gameAccuracyTotal / stat.gameAttempts) : 0;
      var gameParticipation = Math.min(100, stat.gameAttempts * 20);
      var gameScore = hasGame ? Math.round((gameAccuracy * 0.75) + (gameParticipation * 0.25)) : 0;
      var notesScore = hasNotes ? Math.round((stat.notesAccuracy * 0.7) + (stat.notesCompletionPct * 0.3)) : 0;
      var weightNotes = hasNotes ? 0.6 : 0;
      var weightGame = hasGame ? 0.4 : 0;
      var score = (weightNotes + weightGame) > 0
        ? Math.round(((notesScore * weightNotes) + (gameScore * weightGame)) / (weightNotes + weightGame))
        : 0;
      return {
        topic: label,
        label: label,
        score: score,
        notesScore: notesScore,
        gameScore: gameScore,
        notesAccuracy: stat.notesAccuracy,
        notesCompletionPct: stat.notesCompletionPct,
        gameAccuracy: gameAccuracy,
        gameAttempts: stat.gameAttempts,
        bestScore: stat.gameBestScore,
        completed: !!stat.notesCompleted,
        passedSections: stat.passedSections,
        totalSections: stat.totalSections,
        dataPoints: (hasNotes ? 1 : 0) + (hasGame ? 1 : 0)
      };
    }).sort(function(a, b) {
      return b.score - a.score || b.gameAttempts - a.gameAttempts || a.label.localeCompare(b.label);
    });

    var completedTopicLabels = masteryTopics.filter(function(item) { return item.completed; }).map(function(item) { return item.label; });
    var bestScore = entries.reduce(function(maxValue, entry) {
      return Math.max(maxValue, clampNumber_(entry.score, 0));
    }, 0);
    var purchasedIds = Object.keys(purchased || {}).filter(function(key) { return !!purchased[key]; });
    var metrics = {
      totalScore: clampNumber_(student.totalPoints || 0, 0),
      completedChallenges: entries.length,
      completedTopics: completedTopicLabels.length,
      bestScore: bestScore,
      redeemedRoles: purchasedIds.length,
      totalTopics: Number(topicMeta.topics.length || 0),
      completedTopicLabels: completedTopicLabels
    };

    var achievementCatalog = buildAchievementCatalog_(metrics.totalTopics);
    var existingAchievements = student.achievements || {};
    var normalizedAchievements = {};
    var achievements = achievementCatalog.map(function(item) {
      var progress = computeAchievementProgress_(item, metrics);
      var existing = existingAchievements[item.id] || {};
      var unlockedAt = existing.unlockedAt || (progress.unlocked ? Date.now() : 0);
      normalizedAchievements[item.id] = {
        title: item.title,
        description: item.description,
        icon: item.icon,
        unlocked: progress.unlocked,
        unlockedAt: unlockedAt || 0,
        current: progress.current,
        target: progress.target,
        progressPct: progress.progressPct
      };
      return {
        id: item.id,
        title: item.title,
        description: item.description,
        icon: item.icon,
        unlocked: progress.unlocked,
        unlockedAt: unlockedAt || 0,
        current: progress.current,
        target: progress.target,
        progressPct: progress.progressPct
      };
    });

    var unlockedAchievements = achievements.filter(function(item) { return item.unlocked; }).length;
    var strongestTopics = masteryTopics.slice().sort(function(a, b) {
      return b.score - a.score || a.label.localeCompare(b.label);
    }).slice(0, 3);
    var weakestTopics = masteryTopics.filter(function(item) { return item.dataPoints > 0; }).slice().sort(function(a, b) {
      return a.score - b.score || a.label.localeCompare(b.label);
    }).slice(0, 3);
    var radarTopics = masteryTopics.slice().sort(function(a, b) {
      var activityA = (a.dataPoints * 100) + a.gameAttempts + a.notesCompletionPct;
      var activityB = (b.dataPoints * 100) + b.gameAttempts + b.notesCompletionPct;
      return activityB - activityA || b.score - a.score;
    }).slice(0, 6);
    var recommendation = buildProfileRecommendation_(metrics, masteryTopics, topicMeta);
    var ownedFrames = ((student.shop || {}).ownedFrames) || {};
    var ownedFrameCount = Object.keys(ownedFrames).filter(function(key) { return !!ownedFrames[key]; }).length;
    var equippedFrameId = String(((student.shop || {}).equippedFrame) || student.avatarFrameId || '');
    var loginStreak = clampNumber_(student.loginStreakDays || 0, 0);
    var lastLoginDate = String(student.lastLoginDate || '');

    var profile = {
      generatedAt: Date.now(),
      streak: loginStreak,
      lastLoginDate: lastLoginDate,
      equippedFrameId: equippedFrameId,
      metrics: {
        totalScore: metrics.totalScore,
        completedChallenges: metrics.completedChallenges,
        completedTopics: metrics.completedTopics,
        bestScore: metrics.bestScore,
        redeemedRoles: metrics.redeemedRoles,
        ownedFrameCount: ownedFrameCount,
        unlockedAchievements: unlockedAchievements,
        totalAchievements: achievements.length,
        totalTopics: metrics.totalTopics
      },
      achievements: achievements,
      mastery: {
        topics: masteryTopics,
        radarTopics: radarTopics,
        strongestTopics: strongestTopics,
        weakestTopics: weakestTopics
      },
      recommendation: recommendation
    };

    var patchRes = patchFirebase_('students/' + playerKey, {
      achievements: normalizedAchievements,
      profile: profile,
      resultIndexMeta: {
        version: 1,
        updatedAt: Date.now(),
        profileGeneratedAt: profile.generatedAt
      }
    });
    if (patchRes.error) Logger.log('[getStudentProfileData] ' + patchRes.error);
    return profile;
  } catch (e) {
    return { error: '整理學生 profile 時發生錯誤：' + e.message };
  }
}

// ==========================================
// 5. 溫習筆記 - 取得課題列表
// ==========================================
function getNotesTopics(grade, schoolLevel) {
  try {
    var sheet = openFirstSheetById_(getNotesSheetIdForContext_(grade, schoolLevel), '溫習筆記 Sheet ID');
    var data = sheet.getDataRange().getValues();

    var headers = normalizeHeaders_(data.shift());
    var topicIdx = findColumnIndex_(headers, ['課題', 'topic', '主題', 'theme', 'chapter']);
    var gradeIdx = findColumnIndex_(headers, ['級別', 'grade', '年級', 'level', '班級']);
    var sectionIdx = findColumnIndex_(headers, ['section', '章節', '單元']);

    if (topicIdx === -1) return { error: '筆記試算表缺少「課題」欄位。' };

    var seen = {};
    var topics = [];
    for (var i = 0; i < data.length; i++) {
      if (grade && gradeIdx !== -1) {
        var rowGrade = String(data[i][gradeIdx] || "").trim();
        if (rowGrade !== grade) continue;
      }
      var t = String(data[i][topicIdx] || "").trim();
      if (t === "") continue;

      if (!seen[t]) {
        seen[t] = {
          name: t,
          sectionMap: {}
        };
        topics.push(seen[t]);
      }

      if (sectionIdx !== -1) {
        var sectionName = String(data[i][sectionIdx] || '').trim();
        if (sectionName) {
          seen[t].sectionMap[sectionName] = true;
        }
      }
    }
    return topics.map(function(item) {
      return {
        name: item.name,
        sectionCount: Object.keys(item.sectionMap).length
      };
    });
  } catch (e) {
    return { error: '讀取筆記課題錯誤: ' + e.message };
  }
}

// ==========================================
// 6a. 溫習筆記 - 取得指定課題的章節列表
// ==========================================
function getNoteSections(topic, grade, schoolLevel) {
  try {
    var sheet = openFirstSheetById_(getNotesSheetIdForContext_(grade, schoolLevel), '溫習筆記 Sheet ID');
    var data = sheet.getDataRange().getValues();

    var headers = normalizeHeaders_(data.shift());
    var topicIdx   = findColumnIndex_(headers, ['課題', 'topic', '主題', 'theme', 'chapter']);
    var gradeIdx   = findColumnIndex_(headers, ['級別', 'grade', '年級', 'level', '班級']);
    var sectionIdx = findColumnIndex_(headers, ['section', '章節', '單元']);
    var orderIdx   = findColumnIndex_(headers, ['排序', 'order', '順序', '序號']);

    if (topicIdx === -1) return { error: '筆記試算表缺少「課題」欄位。' };
    if (sectionIdx === -1) return { error: '筆記試算表缺少「section」欄位。' };

    var sectionMap = {};
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (String(row[topicIdx] || "").trim().toLowerCase() !== String(topic).toLowerCase()) continue;
      if (grade && gradeIdx !== -1) {
        if (String(row[gradeIdx] || "").trim() !== grade) continue;
      }
      var sec = String(row[sectionIdx] || "").trim();
      if (sec === "") continue;
      if (!sectionMap[sec]) {
        sectionMap[sec] = orderIdx !== -1 ? (Number(row[orderIdx]) || (i + 1)) : (i + 1);
      } else {
        var ord = orderIdx !== -1 ? (Number(row[orderIdx]) || (i + 1)) : (i + 1);
        if (ord < sectionMap[sec]) sectionMap[sec] = ord;
      }
    }

    var sections = Object.keys(sectionMap).map(function(name) {
      return { name: name, order: sectionMap[name] };
    });
    sections.sort(function(a, b) { return a.order - b.order; });

    return sections.map(function(s) { return s.name; });
  } catch (e) {
    return { error: '讀取章節列表錯誤: ' + e.message };
  }
}

// ==========================================
// 6b. 溫習筆記 - 取得指定課題的筆記內容
// ==========================================
function getNoteContent(topic, grade, section, schoolLevel) {
  try {
    var sheet = openFirstSheetById_(getNotesSheetIdForContext_(grade, schoolLevel), '溫習筆記 Sheet ID');
    var data = sheet.getDataRange().getValues();

    var headers = normalizeHeaders_(data.shift());
    var topicIdx   = findColumnIndex_(headers, ['課題', 'topic', '主題', 'theme', 'chapter']);
    var gradeIdx   = findColumnIndex_(headers, ['級別', 'grade', '年級', 'level', '班級']);
    var sectionIdx = findColumnIndex_(headers, ['section', '章節', '單元']);
    var orderIdx   = findColumnIndex_(headers, ['排序', 'order', '順序', '序號']);
    var titleIdx   = findColumnIndex_(headers, ['標題', 'title', '小標題']);
    var contentIdx = findColumnIndex_(headers, ['內容', 'content', '正文', '筆記']);
    var imageIdx   = findColumnIndex_(headers, ['圖片', 'imageurl', 'image', '圖片url', '插圖']);

    if (topicIdx === -1 || contentIdx === -1) {
      return { error: '筆記試算表缺少必要欄位（課題、內容）。' };
    }

    var pages = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowTopic = String(row[topicIdx] || "").trim();
      if (rowTopic.toLowerCase() !== String(topic).toLowerCase()) continue;

      if (grade && gradeIdx !== -1) {
        var rowGrade = String(row[gradeIdx] || "").trim();
        if (rowGrade !== grade) continue;
      }

      if (section && sectionIdx !== -1) {
        var rowSection = String(row[sectionIdx] || "").trim();
        if (rowSection !== section) continue;
      }

      pages.push({
        order: orderIdx !== -1 ? Number(row[orderIdx]) || (i + 1) : (i + 1),
        section: sectionIdx !== -1 ? String(row[sectionIdx] || "").trim() : "",
        title: titleIdx !== -1 ? String(row[titleIdx] || "").trim() : "",
        content: String(row[contentIdx] || "").trim(),
        imageUrl: imageIdx !== -1 ? String(row[imageIdx] || "").trim() : ""
      });
    }

    pages.sort(function(a, b) { return a.order - b.order; });

    if (pages.length === 0) {
      return { error: '找不到「' + topic + '」的筆記內容。' };
    }
    return pages;
  } catch (e) {
    return { error: '讀取筆記內容錯誤: ' + e.message };
  }
}

// ==========================================
// 7. 溫習筆記 - 取得指定課題的測驗題目
// ==========================================
function getNotesQuiz(topic, grade, section, schoolLevel) {
  try {
    var sources = openQuestionTypeSheetsById_(getNotesQuizSheetIdForContext_(grade, schoolLevel), '溫習測驗 Sheet ID');
    var questions = [];

    sources.forEach(function(source) {
      var data = source.sheet.getDataRange().getValues();
      if (!data || data.length < 2) return;

      var headers = normalizeHeaders_(data.shift());
      var findIdx = function(names) { return findColumnIndex_(headers, names); };
      var topicIdx   = findIdx(['課題', 'topic', '主題', 'theme', 'chapter']);
      var gradeIdx   = findIdx(['級別', 'grade', '年級', 'level', '班級']);
      var sectionIdx = findIdx(['section', '章節', '單元']);
      var qIdx       = findIdx(['question', '題目', '問題', 'q']);
      var ansIdx     = findIdx(['answer', '答案', '解答', 'correctanswer', 'ans']);
      var expIdx     = findIdx(['explanation', '詳解', '解釋', '說明', '解析', 'exp']);
      var typeIdx    = findIdx(['questiontype', '題型', '題目類型', 'qtype', 'questionkind']);
      var acceptedIdx = findIdx(['acceptedanswers', 'acceptableanswers', 'altanswers', '可接受答案', '其他答案']);
      var ptsIdx     = findIdx(['積分', 'points', '分數', 'score', 'reward']);
      var imgIdx     = findIdx(['image', 'imageurl', '圖片', '圖片url', '插圖']);
      var matchLeftIdx = findIdx(['matchingleft', 'leftitems', 'leftoptions', 'pairleft', '配對左', '配對左欄', '左欄']);
      var matchRightIdx = findIdx(['matchingright', 'rightitems', 'rightoptions', 'pairright', '配對右', '配對右欄', '右欄']);
      var optAIdx  = findIdx(['optiona', '選項a', 'a', 'option1', '選項1']);
      var optBIdx  = findIdx(['optionb', '選項b', 'b', 'option2', '選項2']);
      var optCIdx  = findIdx(['optionc', '選項c', 'c', 'option3', '選項3']);
      var optDIdx  = findIdx(['optiond', '選項d', 'd', 'option4', '選項4']);

      if (qIdx === -1 || ansIdx === -1) return;

      var optIndices = [];
      if (optAIdx !== -1) optIndices.push(optAIdx);
      if (optBIdx !== -1 && optIndices.indexOf(optBIdx) === -1) optIndices.push(optBIdx);
      if (optCIdx !== -1 && optIndices.indexOf(optCIdx) === -1) optIndices.push(optCIdx);
      if (optDIdx !== -1 && optIndices.indexOf(optDIdx) === -1) optIndices.push(optDIdx);

      if (optIndices.length === 0) {
        var start = Math.min(qIdx, ansIdx) + 1;
        var end = Math.max(qIdx, ansIdx);
        for (var k = start; k < end; k++) optIndices.push(k);
      }

      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rowTopic = topicIdx !== -1 ? String(row[topicIdx] || "").trim() : "";

        if (topicIdx !== -1 && rowTopic.toLowerCase() !== String(topic).toLowerCase()) continue;

        if (grade && gradeIdx !== -1) {
          var rowGrade = String(row[gradeIdx] || "").trim();
          if (rowGrade !== grade) continue;
        }

        if (section && sectionIdx !== -1) {
          var rowSection = String(row[sectionIdx] || "").trim();
          if (rowSection !== section) continue;
        }

        if (String(row[qIdx]).trim() === "") continue;

        var payload = buildQuestionPayload_(row, {
          qIdx: qIdx,
          ansIdx: ansIdx,
          expIdx: expIdx,
          imgIdx: imgIdx,
          typeIdx: typeIdx,
          acceptedIdx: acceptedIdx,
          matchLeftIdx: matchLeftIdx,
          matchRightIdx: matchRightIdx,
          optIndices: optIndices,
          forcedQuestionType: source.questionType,
          defaultExplanation: '',
          pointsValue: ptsIdx !== -1 ? (Number(row[ptsIdx]) || 10) : 10
        });
        if (payload) questions.push(payload);
      }
    });

    if (questions.length === 0) {
      return { error: '找不到「' + topic + '」的測驗題目。' };
    }

    if (section) {
      questions = selectNotesQuizQuestions_(questions, 7);
    } else {
      questions.sort(function() { return Math.random() - 0.5; });
    }
    return questions;
  } catch (e) {
    return { error: '讀取測驗題目錯誤: ' + e.message };
  }
}

function normalizeWrongBookLookupText_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function buildWrongBookImageLookupCacheKey_(source, lookup) {
  lookup = lookup || {};
  return [
    'wbimg',
    String(source || '').trim(),
    String(lookup.question || '').trim(),
    String(lookup.answer || '').trim(),
    String(lookup.topic || '').trim(),
    String(lookup.section || '').trim()
  ].join('::');
}

function getCachedWrongBookImageLookups_(source, lookups) {
  var cache = CacheService.getScriptCache();
  var keys = [];
  var keyToLookup = {};
  (lookups || []).forEach(function(lookup) {
    var cacheKey = buildWrongBookImageLookupCacheKey_(source, lookup);
    keys.push(cacheKey);
    keyToLookup[cacheKey] = lookup;
  });
  if (!keys.length) return { resolved: {}, missing: [] };

  var cachedMap = cache.getAll(keys) || {};
  var resolved = {};
  var missing = [];
  keys.forEach(function(cacheKey) {
    if (Object.prototype.hasOwnProperty.call(cachedMap, cacheKey)) {
      var lookup = keyToLookup[cacheKey];
      if (lookup && cachedMap[cacheKey]) resolved[lookup.key] = cachedMap[cacheKey];
      return;
    }
    missing.push(keyToLookup[cacheKey]);
  });
  return { resolved: resolved, missing: missing };
}

function cacheWrongBookImageLookups_(source, resolvedByKey, lookups) {
  var cache = CacheService.getScriptCache();
  var payload = {};
  (lookups || []).forEach(function(lookup) {
    if (!lookup || !resolvedByKey[lookup.key]) return;
    payload[buildWrongBookImageLookupCacheKey_(source, lookup)] = String(resolvedByKey[lookup.key] || '');
  });
  if (Object.keys(payload).length) cache.putAll(payload, 21600);
}

function buildWrongBookQuestionRowIndex_(rows, config) {
  var index = {};
  var questionIdx = Number(config.questionIdx);
  var answerIdx = Number(config.answerIdx);
  var topicIdx = Number(config.topicIdx);
  var sectionIdx = Number(config.sectionIdx);
  var imageIdx = Number(config.imageIdx);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var questionKey = normalizeWrongBookLookupText_(row[questionIdx]);
    var imageUrl = String(row[imageIdx] || '').trim();
    if (!questionKey || !imageUrl) continue;
    if (!index[questionKey]) index[questionKey] = [];
    index[questionKey].push({
      answer: answerIdx !== -1 ? normalizeWrongBookLookupText_(row[answerIdx]) : '',
      topic: topicIdx !== -1 ? normalizeWrongBookLookupText_(row[topicIdx]) : '',
      section: sectionIdx !== -1 ? normalizeWrongBookLookupText_(row[sectionIdx]) : '',
      imageUrl: imageUrl
    });
  }
  return index;
}

function findWrongBookImageFromCandidates_(lookup, candidates) {
  var rows = Array.isArray(candidates) ? candidates : [];
  for (var i = 0; i < rows.length; i++) {
    var candidate = rows[i] || {};
    if (lookup.answer && candidate.answer && lookup.answer !== candidate.answer) continue;
    if (lookup.topic && candidate.topic && lookup.topic !== candidate.topic) continue;
    if (lookup.section && candidate.section && lookup.section !== candidate.section) continue;
    if (candidate.imageUrl) return candidate.imageUrl;
  }
  return '';
}

function resolveWrongBookImageUrls(items) {
  try {
    var requests = Array.isArray(items) ? items : [];
    if (!requests.length) return {};

    var result = {};
    var pendingGame = {};
    var pendingNotes = {};

    requests.forEach(function(item) {
      item = item || {};
      var key = String(item.key || '').trim();
      var normalizedQuestion = normalizeWrongBookLookupText_(item.question);
      var normalizedAnswer = normalizeWrongBookLookupText_(item.answer);
      if (!key || !normalizedQuestion) return;

      var lookup = {
        key: key,
        question: normalizedQuestion,
        answer: normalizedAnswer,
        topic: normalizeWrongBookLookupText_(item.topic),
        section: normalizeWrongBookLookupText_(item.section)
      };

      if (String(item.source || '').trim() === 'notes') pendingNotes[key] = lookup;
      else pendingGame[key] = lookup;
    });

    var gameLookups = Object.keys(pendingGame).map(function(key) { return pendingGame[key]; });
    var noteLookups = Object.keys(pendingNotes).map(function(key) { return pendingNotes[key]; });
    var cachedGame = getCachedWrongBookImageLookups_('game', gameLookups);
    var cachedNotes = getCachedWrongBookImageLookups_('notes', noteLookups);
    Object.keys(cachedGame.resolved).forEach(function(key) { result[key] = cachedGame.resolved[key]; });
    Object.keys(cachedNotes.resolved).forEach(function(key) { result[key] = cachedNotes.resolved[key]; });

    if (cachedGame.missing.length) {
      var gameSources = openQuestionTypeSheetsById_(getQuestionsSheetId_(), '題庫 Sheet ID');
      var gameIndex = {};
      gameSources.forEach(function(source) {
        var gameData = source.sheet.getDataRange().getValues();
        if (!gameData || gameData.length < 2) return;
        var gameHeaders = normalizeHeaders_(gameData.shift());
        var gameQuestionIdx = findColumnIndex_(gameHeaders, ['question', '題目', '問題', 'q']);
        var gameAnswerIdx = findColumnIndex_(gameHeaders, ['answer', '答案', '解答', 'correctanswer', 'ans']);
        var gameImageIdx = findColumnIndex_(gameHeaders, ['image', 'imageurl', '圖片', '圖片url', '插圖']);
        if (gameQuestionIdx === -1 || gameImageIdx === -1) return;
        var partialIndex = buildWrongBookQuestionRowIndex_(gameData, {
          questionIdx: gameQuestionIdx,
          answerIdx: gameAnswerIdx,
          topicIdx: -1,
          sectionIdx: -1,
          imageIdx: gameImageIdx
        });
        Object.keys(partialIndex).forEach(function(questionKey) {
          gameIndex[questionKey] = (gameIndex[questionKey] || []).concat(partialIndex[questionKey]);
        });
      });
      cachedGame.missing.forEach(function(request) {
        if (result[request.key]) return;
        var found = findWrongBookImageFromCandidates_(request, gameIndex[request.question]);
        if (found) result[request.key] = found;
      });
      cacheWrongBookImageLookups_('game', result, cachedGame.missing);
    }

    if (cachedNotes.missing.length) {
      var noteSources = openQuestionTypeSheetsById_(getNotesQuizSheetId_(), '溫習測驗 Sheet ID');
      var notesIndex = {};
      noteSources.forEach(function(source) {
        var notesData = source.sheet.getDataRange().getValues();
        if (!notesData || notesData.length < 2) return;
        var notesHeaders = normalizeHeaders_(notesData.shift());
        var notesQuestionIdx = findColumnIndex_(notesHeaders, ['question', '題目', '問題', 'q']);
        var notesAnswerIdx = findColumnIndex_(notesHeaders, ['answer', '答案', '解答', 'correctanswer', 'ans']);
        var notesTopicIdx = findColumnIndex_(notesHeaders, ['課題', 'topic', '主題', 'theme', 'chapter']);
        var notesSectionIdx = findColumnIndex_(notesHeaders, ['section', '章節', '單元']);
        var notesImageIdx = findColumnIndex_(notesHeaders, ['image', 'imageurl', '圖片', '圖片url', '插圖']);
        if (notesQuestionIdx === -1 || notesImageIdx === -1) return;
        var partialNotesIndex = buildWrongBookQuestionRowIndex_(notesData, {
          questionIdx: notesQuestionIdx,
          answerIdx: notesAnswerIdx,
          topicIdx: notesTopicIdx,
          sectionIdx: notesSectionIdx,
          imageIdx: notesImageIdx
        });
        Object.keys(partialNotesIndex).forEach(function(questionKey) {
          notesIndex[questionKey] = (notesIndex[questionKey] || []).concat(partialNotesIndex[questionKey]);
        });
      });
      cachedNotes.missing.forEach(function(request) {
        if (result[request.key]) return;
        var found = findWrongBookImageFromCandidates_(request, notesIndex[request.question]);
        if (found) result[request.key] = found;
      });
      cacheWrongBookImageLookups_('notes', result, cachedNotes.missing);
    }

    return result;
  } catch (e) {
    return { error: '補查錯題圖片失敗: ' + e.message };
  }
}

// ==========================================
// API 請求速率限制
// ==========================================
var AI_RATE_LIMIT_PER_MINUTE = 10;
var AI_RATE_LIMIT_PER_DAY = 100;

function checkAiRateLimit_() {
  var cache = CacheService.getScriptCache();
  var now = new Date();
  var minuteKey = 'ai_rpm_' + Utilities.formatDate(now, 'Asia/Hong_Kong', 'yyyyMMddHHmm');
  var dayKey = 'ai_rpd_' + Utilities.formatDate(now, 'Asia/Hong_Kong', 'yyyyMMdd');

  var minuteCount = Number(cache.get(minuteKey) || 0);
  var dayCount = Number(cache.get(dayKey) || 0);

  if (minuteCount >= AI_RATE_LIMIT_PER_MINUTE) {
    return { limited: true, message: '請求太頻繁，請稍後再試。（每分鐘上限 ' + AI_RATE_LIMIT_PER_MINUTE + ' 次）' };
  }
  if (dayCount >= AI_RATE_LIMIT_PER_DAY) {
    return { limited: true, message: '今天的 AI 助手使用次數已達上限（' + AI_RATE_LIMIT_PER_DAY + ' 次），明天再來吧。' };
  }

  cache.put(minuteKey, String(minuteCount + 1), 120);
  cache.put(dayKey, String(dayCount + 1), 86400);
  return { limited: false };
}

// ==========================================
// 8. 數學解難助手 - OpenRouter 代理
// ==========================================
function getOpenRouterApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY') || '';
}

function getOpenRouterPreferredModel_() {
  return PropertiesService.getScriptProperties().getProperty('OPENROUTER_MODEL') || OPENROUTER_MODEL;
}

function getOpenRouterFallbackModel_() {
  return PropertiesService.getScriptProperties().getProperty('OPENROUTER_FALLBACK_MODEL') || OPENROUTER_FALLBACK_MODEL;
}

function getOpenRouterFinalFallbackModel_() {
  return PropertiesService.getScriptProperties().getProperty('OPENROUTER_FINAL_FALLBACK_MODEL') || 'openai/gpt-4o-mini';
}

function authorizeMathHelperAccess() {
  var apiKey = getOpenRouterApiKey_();
  if (!apiKey) {
    return { error: 'Script Properties 尚未設定 OPENROUTER_API_KEY。' };
  }

  var res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/models', {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://script.google.com',
      'X-Title': 'Math Game AI Helper Auth Check'
    },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    return {
      error: 'OpenRouter 連線測試失敗，HTTP ' + code + '：' + sanitizeMathHelperText_(res.getContentText(), 240)
    };
  }

  return { success: true, message: 'OpenRouter 授權與外部連線測試成功。' };
}

function sanitizeMathHelperText_(value, maxLength) {
  var text = String(value == null ? '' : value);
  if (!maxLength || text.length <= maxLength) return text;
  return text.slice(0, maxLength);
}

function buildMathHelperSystemPrompt_() {
  return [
    '你是一位親切、專業、擅長引導學生思考的「數學解難助手」。',
    '請使用繁體中文（香港常用語氣）回答。',
    '核心原則：',
    '1. 不要一開始直接給最終答案。',
    '2. 先用蘇格拉底式提問、提示、拆步驟來引導學生。',
    '3. 如果學生已經做錯，你可以指出思路哪裡偏了，但不要直接把正解整句端出來。',
    '4. 如果系統提供了「老師詳解」或「正確答案相關背景」，那只是給你作隱藏參考，不要直接照抄給學生。',
    '5. 當你判斷學生已大致掌握，可用以下格式整理：',
    '---',
    '### ✨ 溫習筆記：[題目關鍵字]',
    '🔢 **運算步驟**：',
    '1. ...',
    '🧠 **邏輯解釋**：',
    '...',
    '---',
    '### 🎯 隨堂挑戰：',
    '[出一條相關的小挑戰題]',
    '---',
    '請盡量精準、簡潔、可操作；需要公式時可用 KaTeX 語法。'
  ].join('\n');
}

function buildMathHelperContextPrompt_(context) {
  context = context || {};
  var lines = [];
  if (context.source) lines.push('來源：' + sanitizeMathHelperText_(context.source, 80));
  if (context.question) lines.push('題目：' + sanitizeMathHelperText_(context.question, 4000));
  if (context.studentAnswer) lines.push('學生先前答案：' + sanitizeMathHelperText_(context.studentAnswer, 1200));
  if (context.explanation) lines.push('老師詳解（隱藏背景，只供你判斷，不要直接照抄）：' + sanitizeMathHelperText_(context.explanation, 5000));
  if (context.topic) lines.push('課題：' + sanitizeMathHelperText_(context.topic, 120));
  if (context.section) lines.push('章節：' + sanitizeMathHelperText_(context.section, 120));
  if (!lines.length) return '';
  return '以下是本輪對話的隱藏背景資料：\n' + lines.join('\n');
}

function normalizeMathHelperContent_(content) {
  if (typeof content === 'string') {
    return sanitizeMathHelperText_(content, 12000);
  }
  if (!Array.isArray(content)) return '';
  var normalized = [];
  for (var i = 0; i < content.length; i++) {
    var item = content[i] || {};
    if (item.type === 'text') {
      normalized.push({ type: 'text', text: sanitizeMathHelperText_(item.text, 12000) });
      continue;
    }
    if (item.type === 'image_url' && item.image_url && item.image_url.url) {
      var url = String(item.image_url.url || '');
      if (url.length > 4500000) continue;
      normalized.push({ type: 'image_url', image_url: { url: url } });
    }
  }
  return normalized.length ? normalized : '';
}

function normalizeMathHelperMessages_(messages) {
  var arr = Array.isArray(messages) ? messages : [];
  var normalized = [];
  var start = Math.max(0, arr.length - 12);
  for (var i = start; i < arr.length; i++) {
    var msg = arr[i] || {};
    var role = msg.role === 'assistant' ? 'assistant' : 'user';
    var content = normalizeMathHelperContent_(msg.content);
    if (!content || (Array.isArray(content) && !content.length)) continue;
    normalized.push({ role: role, content: content });
  }
  return normalized;
}

function stripImagesFromMathHelperMessages_(messages) {
  var arr = Array.isArray(messages) ? messages : [];
  return arr.map(function(msg) {
    if (!msg || !Array.isArray(msg.content)) return msg;
    var textParts = msg.content.filter(function(item) {
      return item && item.type === 'text' && String(item.text || '').trim() !== '';
    }).map(function(item) {
      return String(item.text || '');
    });
    return {
      role: msg.role,
      content: textParts.join('\n') || '請協助我理解這條數學題。'
    };
  });
}

function extractMathHelperReply_(body) {
  if (!body || !body.choices || !body.choices.length || !body.choices[0].message) return '';
  var content = body.choices[0].message.content;
  if (typeof content === 'string') return content;
  if (!Array.isArray(content)) return '';
  return content.map(function(part) {
    return part && part.type === 'text' ? String(part.text || '') : '';
  }).join('');
}

function buildOpenRouterHeaders_(apiKey, title) {
  return {
    Authorization: 'Bearer ' + apiKey,
    'HTTP-Referer': 'https://script.google.com',
    'X-Title': title || 'Math Game AI Helper'
  };
}

function callOpenRouter_(apiKey, payload, title) {
  var res = UrlFetchApp.fetch(OPENROUTER_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: buildOpenRouterHeaders_(apiKey, title),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var raw = res.getContentText();
  var body = {};
  try {
    body = JSON.parse(raw || '{}');
  } catch (parseErr) {}

  return {
    code: code,
    raw: raw,
    body: body
  };
}

function buildOpenRouterErrorMessage_(response) {
  var body = response && response.body ? response.body : {};
  var raw = response && response.raw ? response.raw : '';
  var message = body && body.error && body.error.message ? String(body.error.message) : String(raw || '未知錯誤');
  var metadata = body && body.error && body.error.metadata ? body.error.metadata : null;
  var provider = metadata && metadata.provider_name ? '（provider: ' + metadata.provider_name + '）' : '';
  var lower = message.toLowerCase();

  if (lower.indexOf('violation of provider terms of service') !== -1) {
    return 'AI 服務已收到請求，但上游模型供應商拒絕處理這次內容' + provider + '。這通常不是 Google 授權問題，而是目前選用的模型或輸入內容被供應商政策攔截。你可以先重試純文字題目，或改用另一個 OpenRouter 模型。';
  }

  return 'AI 服務暫時未能回應：' + sanitizeMathHelperText_(message, 240);
}

function isProviderTermsError_(response) {
  var message = response && response.body && response.body.error && response.body.error.message
    ? String(response.body.error.message)
    : String((response && response.raw) || '');
  return message.toLowerCase().indexOf('violation of provider terms of service') !== -1;
}

function tryOpenRouterMathHelper_(apiKey, payloadMessages) {
  var attemptedModels = [];
  var modelsToTry = [
    getOpenRouterPreferredModel_(),
    getOpenRouterFallbackModel_(),
    getOpenRouterFinalFallbackModel_()
  ].filter(function(model, index, arr) {
    return !!model && arr.indexOf(model) === index;
  });

  var lastResponse = null;
  var usingTextOnly = false;

  for (var phase = 0; phase < 2; phase++) {
    var phaseMessages = phase === 0 ? payloadMessages : stripImagesFromMathHelperMessages_(payloadMessages);
    usingTextOnly = phase === 1;

    for (var i = 0; i < modelsToTry.length; i++) {
      var model = modelsToTry[i];
      attemptedModels.push((usingTextOnly ? 'text-only:' : 'full:') + model);
      var response = callOpenRouter_(apiKey, {
        model: model,
        messages: phaseMessages,
        temperature: 0.55,
        max_tokens: 900
      }, usingTextOnly ? 'Math Game AI Helper Text Only' : 'Math Game AI Helper');

      if (response.code >= 200 && response.code < 300) {
        return {
          response: response,
          model: model,
          textOnly: usingTextOnly,
          attemptedModels: attemptedModels
        };
      }

      lastResponse = response;

      if (!isProviderTermsError_(response)) {
        break;
      }
    }
  }

  return {
    error: buildOpenRouterErrorMessage_(lastResponse),
    attemptedModels: attemptedModels,
    lastResponse: lastResponse
  };
}

function solveMathWithAI(request) {
  try {
    var rateCheck = checkAiRateLimit_();
    if (rateCheck.limited) {
      return { error: rateCheck.message };
    }

    var apiKey = getOpenRouterApiKey_();
    if (!apiKey) {
      return { error: '數學解難助手尚未完成啟用。請先在 Apps Script Script Properties 設定 OPENROUTER_API_KEY。' };
    }

    request = request || {};
    var messages = normalizeMathHelperMessages_(request.messages);
    if (!messages.length) {
      return { error: '未收到可用的提問內容。' };
    }

    var payloadMessages = [
      { role: 'system', content: buildMathHelperSystemPrompt_() }
    ];

    var contextPrompt = buildMathHelperContextPrompt_(request.context);
    if (contextPrompt) payloadMessages.push({ role: 'system', content: contextPrompt });

    payloadMessages = payloadMessages.concat(messages);

    var attempt = tryOpenRouterMathHelper_(apiKey, payloadMessages);
    if (attempt.error) {
      return { error: attempt.error };
    }

    var response = attempt.response;
    var reply = extractMathHelperReply_(response.body).trim();
    if (!reply) {
      return { error: 'AI 服務未返回有效內容。' };
    }

    return {
      success: true,
      reply: reply,
      usage: response.body.usage || null,
      model: response.body.model || attempt.model || 'openrouter-default',
      textOnlyFallback: !!attempt.textOnly
    };
  } catch (e) {
    return { error: 'AI 助手發生系統錯誤：' + e.message };
  }
}

// ==========================================
// 9. 房間自動清理（可設 time-based trigger 定期執行）
// ==========================================
function cleanupStaleRooms() {
  try {
    var maxAgeMs = 3600000; // 1 hour
    var now = Date.now();
    var url = getFirebaseDbUrl_() + '/rooms.json?shallow=true';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return;
    var roomIds = JSON.parse(res.getContentText() || '{}');
    if (!roomIds) return;

    var cleaned = 0;
    Object.keys(roomIds).forEach(function(roomId) {
      var roomUrl = getFirebaseDbUrl_() + '/rooms/' + roomId + '.json';
      var roomRes = UrlFetchApp.fetch(roomUrl, { muteHttpExceptions: true });
      if (roomRes.getResponseCode() !== 200) return;
      var room = JSON.parse(roomRes.getContentText() || '{}');
      if (!room || !room.createdAt) return;

      var age = now - Number(room.createdAt);
      if (age > maxAgeMs || room.status === 'host_left' || room.status === 'closed') {
        UrlFetchApp.fetch(roomUrl, { method: 'delete', muteHttpExceptions: true });
        cleaned++;
      }
      // Also clean up host_disconnected rooms past their resume deadline
      if (room.status === 'host_disconnected' && room.resumeDeadline && now > room.resumeDeadline) {
        UrlFetchApp.fetch(roomUrl, { method: 'delete', muteHttpExceptions: true });
        cleaned++;
      }
    });

    Logger.log('cleanupStaleRooms: removed ' + cleaned + ' rooms');
    return { cleaned: cleaned };
  } catch (e) {
    Logger.log('cleanupStaleRooms error: ' + e.message);
    return { error: e.message };
  }
}

// ==========================================
// 進行中工作階段持久化 (Active Session Persistence)
// ==========================================
function saveActiveSession(playerInfo, sessionType, sessionData) {
  try {
    var playerKey = buildStudentResultPlayerKey_(playerInfo);
    if (!playerKey) return { error: '缺少學生資料。' };
    var type = String(sessionType || 'challenge').trim();
    if (type !== 'challenge' && type !== 'notes') return { error: '不支援的 session 類型。' };
    var data = sessionData || {};
    data.updatedAt = Date.now();
    return patchFirebase_('students/' + playerKey + '/activeSessions/' + type, data);
  } catch (e) {
    return { error: e.message };
  }
}

function getActiveSession(playerInfo, sessionType) {
  try {
    var playerKey = buildStudentResultPlayerKey_(playerInfo);
    if (!playerKey) return { error: '缺少學生資料。' };
    var type = String(sessionType || 'challenge').trim();
    return fetchFromFirebase_('students/' + playerKey + '/activeSessions/' + type);
  } catch (e) {
    return { error: e.message };
  }
}

function clearActiveSession(playerInfo, sessionType) {
  try {
    var playerKey = buildStudentResultPlayerKey_(playerInfo);
    if (!playerKey) return { error: '缺少學生資料。' };
    var type = String(sessionType || 'challenge').trim();
    var cleanPath = 'students/' + playerKey + '/activeSessions/' + type;
    var url = getFirebaseDbUrl_() + '/' + cleanPath + '.json';
    UrlFetchApp.fetch(url, { method: 'delete', muteHttpExceptions: true });
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}
