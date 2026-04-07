// --- 設定區 ---
const QUESTIONS_SHEET_ID = "1RpKFnPmPDvDvW__nwzZWBnB5jlO6T_aHR9LellTNNiM";
const NOTES_SHEET_ID = "1i-4_v1u9Q7yhTXESEP4oDdeHCuDEsEpicb3RrUd0oks";
const NOTES_QUIZ_SHEET_ID = "1hDpql-EG8zne7NDaabjfuy6Ip1nAxmP-H5I2woEX-tw";
const FIREBASE_DB_URL = "https://math-game-3747d-default-rtdb.firebaseio.com";
const RESULTS_SPREADSHEET_ID = "1Ws1Q3uMnpr4erM3vBNbzL7j1a035rNb3pEAUCAqEmO0";
const OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions";
const OPENROUTER_MODEL = "google/gemini-2.5-flash";
const OPENROUTER_FALLBACK_MODEL = "meta-llama/llama-3.3-70b-instruct";

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
    var url = FIREBASE_DB_URL + '/' + cleanPath + '.json';
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

function formatResultTime_() {
  var tz = Session.getScriptTimeZone() || 'Asia/Hong_Kong';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
}

function saveToSheet(data, type) {
  try {
    var t = String(type || '').toLowerCase();
    var sheetName = t === 'multi' ? 'results(multiplayers)' : 'results(single player)';
    var ss = SpreadsheetApp.openById(RESULTS_SPREADSHEET_ID);
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
    const ss = SpreadsheetApp.openById(QUESTIONS_SHEET_ID);
    const sheet = ss.getSheets()[0];
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
    const ss = SpreadsheetApp.openById(QUESTIONS_SHEET_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    const headers = normalizeHeaders_(data.shift());
    const findIdx = function(names) { return findColumnIndex_(headers, names); };

    const themeIdx = findIdx(['theme', '主題', '類型', 'category', 'topic']);
    const diffIdx  = findIdx(['difficulty', '難度', 'level']);
    const gradeIdx = findIdx(['級別', 'grade', '年級']);
    const qIdx     = findIdx(['question', '題目', '問題', 'q']);
    const ansIdx   = findIdx(['answer', '答案', '解答', 'correctanswer', 'ans']);
    const expIdx   = findIdx(['explanation', '詳解', '解釋', '說明', '解析', 'exp']);

    const imgIdx   = findIdx(['image', 'imageurl', '圖片', '圖片url', '插圖']);

    const optAIdx  = findIdx(['optiona', '選項a', 'a', 'option1', '選項1', 'choicea', 'choice1', '選擇a', '選擇1', 'opt1', 'opta', '1', '選項', '選擇', 'option', 'options', 'choices']);
    const optBIdx  = findIdx(['optionb', '選項b', 'b', 'option2', '選項2', 'choiceb', 'choice2', '選擇b', '選擇2', 'opt2', 'optb', '2']);
    const optCIdx  = findIdx(['optionc', '選項c', 'c', 'option3', '選項3', 'choicec', 'choice3', '選擇c', '選擇3', 'opt3', 'optc', '3']);
    const optDIdx  = findIdx(['optiond', '選項d', 'd', 'option4', '選項4', 'choiced', 'choice4', '選擇d', '選擇4', 'opt4', 'optd', '4']);

    if (qIdx === -1 || ansIdx === -1) {
      return { error: '無法讀取題庫：缺少「題目」或「答案」欄位。' };
    }

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

    const difficultyMap = { '簡單': 'easy', '普通': 'medium', '中等': 'medium', '困難': 'hard' };
    const targetDiff = difficultyMap[difficulty] || difficulty;

    let filtered = data.filter(row => {
      const rowDiff = diffIdx !== -1 ? String(row[diffIdx] || "").trim().toLowerCase() : "";
      const rowTheme = themeIdx !== -1 ? String(row[themeIdx] || "").trim() : "";
      const rowGrade = gradeIdx !== -1 ? String(row[gradeIdx] || "").trim() : "";
      const diffMatch = (!targetDiff || targetDiff === '全部') ? true : (rowDiff === targetDiff);
      const themeMatch = (!theme || theme === '全部') ? true : (rowTheme.toLowerCase() === String(theme).toLowerCase());
      const gradeMatch = (!grade || grade === '全部') ? true : (rowGrade === grade);
      const isNotEmpty = String(row[qIdx]).trim() !== "";
      return diffMatch && themeMatch && gradeMatch && isNotEmpty;
    });

    if (filtered.length === 0) {
      return { error: '找不到符合條件的題目 (難度: ' + difficulty + ', 主題: ' + theme + ')' };
    }

    filtered.sort(() => Math.random() - 0.5);
    filtered = filtered.slice(0, 10);

    const questions = filtered.map(row => {
      let options = [];
      optIndices.forEach(idx => {
        if (idx < row.length && String(row[idx]).trim() !== "") {
          options.push(String(row[idx]).trim());
        }
      });

      if (options.length === 1) {
        let singleOpt = options[0];
        if (singleOpt.includes(',') || singleOpt.includes('，')) {
          options = singleOpt.split(/[,，]/).map(s => s.trim()).filter(s => s !== "");
        } else if (singleOpt.includes('\n')) {
          options = singleOpt.split('\n').map(s => s.trim()).filter(s => s !== "");
        } else if (/A[\.\)]/i.test(singleOpt) && /B[\.\)]/i.test(singleOpt)) {
          let temp = singleOpt.replace(/A[\.\)]/gi, '||A.').replace(/B[\.\)]/gi, '||B.').replace(/C[\.\)]/gi, '||C.').replace(/D[\.\)]/gi, '||D.');
          options = temp.split('||').map(s => s.trim()).filter(s => s !== "");
        } else {
          let spaceSplit = singleOpt.split(/\s+/).map(s => s.trim()).filter(s => s !== "");
          if (spaceSplit.length > 1) options = spaceSplit;
        }
      }
      if (options.length === 0) options = ["選項讀取失敗"];

      return {
        question: String(row[qIdx]).trim(),
        options: options,
        answer: String(row[ansIdx]).trim(),
        explanation: (expIdx !== -1 && String(row[expIdx]).trim() !== "") ? String(row[expIdx]).trim() : "請留意計算步驟喔！",
        imageUrl: (imgIdx !== -1 && String(row[imgIdx] || "").trim() !== "") ? String(row[imgIdx]).trim() : ""
      };
    });

    return questions;
  } catch (e) {
    return { error: '系統錯誤: ' + e.message };
  }
}

// ==========================================
// 3. 儲存單人模式成績 (Firebase)
// ==========================================
function saveSingleResult(playerInfo, score, meta) {
  try {
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
    var sheetRes = saveToSheet(payload, 'multi');
    if (sheetRes && sheetRes.error) {
      Logger.log('[finishGame] Sheet 寫入失敗（已忽略，不影響 Firebase）: ' + sheetRes.error);
    }
    return fbRes;
  } catch (e) {
    return { error: e.message };
  }
}

// ==========================================
// 5. 溫習筆記 - 取得課題列表
// ==========================================
function getNotesTopics(grade) {
  try {
    var ss = SpreadsheetApp.openById(NOTES_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    var headers = normalizeHeaders_(data.shift());
    var topicIdx = findColumnIndex_(headers, ['課題', 'topic', '主題', 'theme', 'chapter']);
    var gradeIdx = findColumnIndex_(headers, ['級別', 'grade', '年級', 'level', '班級']);

    if (topicIdx === -1) return { error: '筆記試算表缺少「課題」欄位。' };

    var seen = {};
    var topics = [];
    for (var i = 0; i < data.length; i++) {
      if (grade && gradeIdx !== -1) {
        var rowGrade = String(data[i][gradeIdx] || "").trim();
        if (rowGrade !== grade) continue;
      }
      var t = String(data[i][topicIdx] || "").trim();
      if (t !== "" && !seen[t]) {
        seen[t] = true;
        topics.push(t);
      }
    }
    return topics;
  } catch (e) {
    return { error: '讀取筆記課題錯誤: ' + e.message };
  }
}

// ==========================================
// 6a. 溫習筆記 - 取得指定課題的章節列表
// ==========================================
function getNoteSections(topic, grade) {
  try {
    var ss = SpreadsheetApp.openById(NOTES_SHEET_ID);
    var sheet = ss.getSheets()[0];
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
function getNoteContent(topic, grade, section) {
  try {
    var ss = SpreadsheetApp.openById(NOTES_SHEET_ID);
    var sheet = ss.getSheets()[0];
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
function getNotesQuiz(topic, grade, section) {
  try {
    var ss = SpreadsheetApp.openById(NOTES_QUIZ_SHEET_ID);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    var headers = normalizeHeaders_(data.shift());
    var findIdx = function(names) { return findColumnIndex_(headers, names); };

    var topicIdx   = findIdx(['課題', 'topic', '主題', 'theme', 'chapter']);
    var gradeIdx   = findIdx(['級別', 'grade', '年級', 'level', '班級']);
    var sectionIdx = findIdx(['section', '章節', '單元']);
    var qIdx       = findIdx(['question', '題目', '問題', 'q']);
    var ansIdx     = findIdx(['answer', '答案', '解答', 'correctanswer', 'ans']);
    var expIdx     = findIdx(['explanation', '詳解', '解釋', '說明', '解析', 'exp']);
    var ptsIdx     = findIdx(['積分', 'points', '分數', 'score', 'reward']);
    var imgIdx     = findIdx(['image', 'imageurl', '圖片', '圖片url', '插圖']);

    var optAIdx  = findIdx(['optiona', '選項a', 'a', 'option1', '選項1']);
    var optBIdx  = findIdx(['optionb', '選項b', 'b', 'option2', '選項2']);
    var optCIdx  = findIdx(['optionc', '選項c', 'c', 'option3', '選項3']);
    var optDIdx  = findIdx(['optiond', '選項d', 'd', 'option4', '選項4']);

    if (qIdx === -1 || ansIdx === -1) {
      return { error: '測驗題庫缺少「題目」或「答案」欄位。' };
    }

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

    var questions = [];
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

      var options = [];
      optIndices.forEach(function(idx) {
        if (idx < row.length && String(row[idx]).trim() !== "") {
          options.push(String(row[idx]).trim());
        }
      });
      if (options.length === 0) options = ["選項讀取失敗"];

      questions.push({
        question: String(row[qIdx]).trim(),
        options: options,
        answer: String(row[ansIdx]).trim(),
        explanation: (expIdx !== -1 && String(row[expIdx]).trim() !== "") ? String(row[expIdx]).trim() : "",
        points: ptsIdx !== -1 ? (Number(row[ptsIdx]) || 10) : 10,
        imageUrl: (imgIdx !== -1 && String(row[imgIdx] || "").trim() !== "") ? String(row[imgIdx]).trim() : ""
      });
    }

    if (questions.length === 0) {
      return { error: '找不到「' + topic + '」的測驗題目。' };
    }

    questions.sort(function() { return Math.random() - 0.5; });
    if (section) {
      questions = questions.slice(0, 5);
    }
    return questions;
  } catch (e) {
    return { error: '讀取測驗題目錯誤: ' + e.message };
  }
}

function normalizeWrongBookLookupText_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
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

    if (Object.keys(pendingGame).length) {
      var gameData = SpreadsheetApp.openById(QUESTIONS_SHEET_ID).getSheets()[0].getDataRange().getValues();
      var gameHeaders = normalizeHeaders_(gameData.shift());
      var gameQuestionIdx = findColumnIndex_(gameHeaders, ['question', '題目', '問題', 'q']);
      var gameAnswerIdx = findColumnIndex_(gameHeaders, ['answer', '答案', '解答', 'correctanswer', 'ans']);
      var gameImageIdx = findColumnIndex_(gameHeaders, ['image', 'imageurl', '圖片', '圖片url', '插圖']);
      if (gameQuestionIdx !== -1 && gameImageIdx !== -1) {
        for (var i = 0; i < gameData.length; i++) {
          var row = gameData[i];
          var questionKey = normalizeWrongBookLookupText_(row[gameQuestionIdx]);
          var answerKey = gameAnswerIdx !== -1 ? normalizeWrongBookLookupText_(row[gameAnswerIdx]) : '';
          var imageUrl = String(row[gameImageIdx] || '').trim();
          if (!questionKey || !imageUrl) continue;

          Object.keys(pendingGame).forEach(function(requestKey) {
            var request = pendingGame[requestKey];
            if (result[request.key]) return;
            if (request.question !== questionKey) return;
            if (request.answer && answerKey && request.answer !== answerKey) return;
            result[request.key] = imageUrl;
          });
        }
      }
    }

    if (Object.keys(pendingNotes).length) {
      var notesData = SpreadsheetApp.openById(NOTES_QUIZ_SHEET_ID).getSheets()[0].getDataRange().getValues();
      var notesHeaders = normalizeHeaders_(notesData.shift());
      var notesQuestionIdx = findColumnIndex_(notesHeaders, ['question', '題目', '問題', 'q']);
      var notesAnswerIdx = findColumnIndex_(notesHeaders, ['answer', '答案', '解答', 'correctanswer', 'ans']);
      var notesTopicIdx = findColumnIndex_(notesHeaders, ['課題', 'topic', '主題', 'theme', 'chapter']);
      var notesSectionIdx = findColumnIndex_(notesHeaders, ['section', '章節', '單元']);
      var notesImageIdx = findColumnIndex_(notesHeaders, ['image', 'imageurl', '圖片', '圖片url', '插圖']);
      if (notesQuestionIdx !== -1 && notesImageIdx !== -1) {
        for (var j = 0; j < notesData.length; j++) {
          var noteRow = notesData[j];
          var noteQuestionKey = normalizeWrongBookLookupText_(noteRow[notesQuestionIdx]);
          var noteAnswerKey = notesAnswerIdx !== -1 ? normalizeWrongBookLookupText_(noteRow[notesAnswerIdx]) : '';
          var noteTopicKey = notesTopicIdx !== -1 ? normalizeWrongBookLookupText_(noteRow[notesTopicIdx]) : '';
          var noteSectionKey = notesSectionIdx !== -1 ? normalizeWrongBookLookupText_(noteRow[notesSectionIdx]) : '';
          var noteImageUrl = String(noteRow[notesImageIdx] || '').trim();
          if (!noteQuestionKey || !noteImageUrl) continue;

          Object.keys(pendingNotes).forEach(function(requestKey) {
            var request = pendingNotes[requestKey];
            if (result[request.key]) return;
            if (request.question !== noteQuestionKey) return;
            if (request.answer && noteAnswerKey && request.answer !== noteAnswerKey) return;
            if (request.topic && noteTopicKey && request.topic !== noteTopicKey) return;
            if (request.section && noteSectionKey && request.section !== noteSectionKey) return;
            result[request.key] = noteImageUrl;
          });
        }
      }
    }

    return result;
  } catch (e) {
    return { error: '補查錯題圖片失敗: ' + e.message };
  }
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
