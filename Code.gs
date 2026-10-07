/**
 * Google 試算表雲端排行榜 — 銀河小蜜蜂
 *
 * 設定：
 * 1. 建立 Google 試算表 → 擴充功能 → Apps Script，貼上本檔。
 * 2. 填寫下方 SPREADSHEET_ID：試算表網址 /d/ 與 /edit 之間的文字。
 * 3. 在編輯器執行 setupLeaderboard，完成 Google 授權。
 * 4. 部署 → 新增部署 → 網頁應用程式。
 *    執行身分：我；誰可以存取：所有人（含未登入者）。
 *    若帳號政策沒有此選項，公開遊戲無法使用這個部署。
 * 5. 網頁使用部署的 /exec 網址；更新程式後需更新部署版本。
 *
 * 讀榜 GET：/exec?action=top
 * 寫榜 POST：{"action":"submit","nickname":"MuMu","score":1500}
 * 成功：{ok:true, mode:"score", leaderboard:[{rank,nickname,score,submittedAt}]}
 * 寫榜另含 saved、personalBest。
 * 失敗：{ok:false,error:{code,message}}；前端必須檢查 ok。
 *
 * 網頁串接（放在遊戲原有 IIFE 內才能取得 score）：
 * const LEADERBOARD_URL = '貼上部署的 /exec 網址';
 * async function submitScore(nickname, finalScore) {
 *   const response = await fetch(LEADERBOARD_URL, {
 *     method: 'POST',
 *     headers: { 'Content-Type': 'text/plain;charset=utf-8' },
 *     body: JSON.stringify({action:'submit', nickname, score:finalScore}),
 *     redirect: 'follow', credentials: 'omit'
 *   });
 *   const data = await response.json();
 *   if (!data.ok) throw new Error(data.error.message);
 *   return data;
 * }
 * async function loadTop10() {
 *   const response = await fetch(LEADERBOARD_URL + '?action=top', {
 *     redirect:'follow', credentials:'omit'
 *   });
 *   const data = await response.json();
 *   if (!data.ok) throw new Error(data.error.message);
 *   return data.leaderboard;
 * }
 * 在 endGame() 中取得玩家暱稱，再呼叫 submitScore(nickname, score)。
 * 顯示暱稱時使用 textContent，避免把玩家輸入當成 HTML。
 * 使用 text/plain 可避免 application/json 引發的跨網域預檢。
 * 不要使用 mode:'no-cors'，否則無法讀取成功回應。
 *
 * 排名規則：每個暱稱一筆最佳成績；同分先達成者優先。
 * 暱稱忽略大小寫。同名玩家會共用成績，暱稱並非登入驗證。
 * 此公開 API 接受前端傳入的成績，適用休閒遊戲，無法防止偽造成績。
 * 此檔提供後端；尚未把暱稱輸入框與榜單介面加入 HTML。
 */

const CONFIG = Object.freeze({
  SPREADSHEET_ID: '請填入你的Google試算表ID',
  MODE: 'score', // 'score'：高分優先；'time'：時間短優先
  SHEET_PREFIX: 'Leaderboard_', // score/time 各自一張工作表
  TOP_LIMIT: 10,
  MAX_VALUE: 1000000000, // 可依遊戲調整；只能限制範圍，不能驗證真實性
  MAX_NICKNAME_LENGTH: 20
});

const HEADERS = ['nickname', 'score', 'submittedAt'];

function setupLeaderboard() {
  return withLock_(function () {
    const sheet = getSheet_();
    return { sheet: sheet.getName(), mode: CONFIG.MODE };
  });
}

function doGet(e) {
  return respond_(function () {
    const action = (e && e.parameter && e.parameter.action) || 'top';
    if (action !== 'top') fail_('INVALID_ACTION', 'GET 僅支援 action=top。');
    return withLock_(function () {
      return { ok: true, mode: CONFIG.MODE, leaderboard: top10_(getSheet_()) };
    });
  });
}

function doPost(e) {
  return respond_(function () {
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > 4096) fail_('INVALID_BODY', '請傳送有效的 JSON 資料。');
    let data;
    try { data = JSON.parse(raw); }
    catch (_) { fail_('INVALID_JSON', 'JSON 格式錯誤。'); }
    if (!data || Array.isArray(data) || typeof data !== 'object') {
      fail_('INVALID_BODY', '資料必須是 JSON 物件。');
    }
    if (data.action !== 'submit') fail_('INVALID_ACTION', 'POST 僅支援 action=submit。');
    const nickname = validateNickname_(data.nickname);
    const value = data.score;
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 ||
        value > CONFIG.MAX_VALUE || (CONFIG.MODE === 'score' && !Number.isSafeInteger(value)) ||
        (CONFIG.MODE === 'time' && value === 0)) {
      fail_('INVALID_SCORE', CONFIG.MODE === 'time'
        ? '時間必須是大於 0 的有效數字，單位為秒。'
        : '分數必須是範圍內的非負整數。');
    }
    return withLock_(function () {
      const sheet = getSheet_();
      const records = readRecords_(sheet);
      const previous = records.find(function (r) {
        return r.nickname.toLowerCase() === nickname.toLowerCase();
      });
      const saved = !previous || better_(value, previous.score);
      if (saved) {
        const row = previous ? previous.row : sheet.getLastRow() + 1;
        sheet.getRange(row, 1, 1, 3).setValues([[nickname, value, new Date()]]);
        SpreadsheetApp.flush();
      }
      return {
        ok: true, mode: CONFIG.MODE, saved: saved,
        personalBest: saved ? value : previous.score,
        leaderboard: top10_(sheet)
      };
    });
  });
}

function validateNickname_(input) {
  if (typeof input !== 'string') fail_('INVALID_NICKNAME', '請輸入玩家暱稱。');
  const name = input.normalize('NFKC').trim();
  if (!name || Array.from(name).length > CONFIG.MAX_NICKNAME_LENGTH ||
      /[\u0000-\u001f\u007f-\u009f\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]/.test(name) ||
      /^[=+@'\-]/.test(name)) {
    fail_('INVALID_NICKNAME', '暱稱需為 1–20 字，不可含控制字元或以 =、+、-、@、單引號開頭。');
  }
  return name;
}

function getSheet_() {
  if (!['score', 'time'].includes(CONFIG.MODE)) fail_('CONFIG_ERROR', 'MODE 請設為 score 或 time。');
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.indexOf('請填入') !== -1) {
    fail_('CONFIG_ERROR', '請先填寫 SPREADSHEET_ID。');
  }
  const book = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const name = CONFIG.SHEET_PREFIX + CONFIG.MODE;
  let sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  } else {
    const actual = sheet.getRange(1, 1, 1, 3).getValues()[0];
    if (actual.join('|') !== HEADERS.join('|')) {
      fail_('SHEET_FORMAT', '排行榜欄位不符，請使用獨立工作表並保留原有標題列。');
    }
  }
  return sheet;
}

function readRecords_(sheet) {
  const count = sheet.getLastRow() - 1;
  if (count <= 0) return [];
  return sheet.getRange(2, 1, count, 3).getValues().map(function (row, index) {
    return { row: index + 2, nickname: String(row[0]), score: row[1], date: row[2] };
  }).filter(function (r) {
    return r.nickname.trim() && typeof r.score === 'number' && Number.isFinite(r.score) &&
      r.score >= 0 && r.date instanceof Date && Number.isFinite(r.date.getTime());
  });
}

function better_(a, b) { return CONFIG.MODE === 'time' ? a < b : a > b; }

function top10_(sheet) {
  const records = readRecords_(sheet).sort(function (a, b) {
    const difference = CONFIG.MODE === 'time' ? a.score - b.score : b.score - a.score;
    return difference || a.date.getTime() - b.date.getTime() || a.row - b.row;
  });
  const seen = new Set();
  return records.filter(function (r) {
    const key = r.nickname.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).slice(0, CONFIG.TOP_LIMIT).map(function (r, index) {
    return { rank: index + 1, nickname: r.nickname, score: r.score, submittedAt: r.date.toISOString() };
  });
}

function withLock_(task) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) fail_('BUSY', '排行榜忙碌中，請稍後重試。');
  try { return task(); }
  finally { lock.releaseLock(); }
}

function fail_(code, message) {
  const error = new Error(message);
  error.publicCode = code;
  throw error;
}

function respond_(task) {
  let result;
  try { result = task(); }
  catch (error) {
    console.error(error);
    result = { ok: false, error: {
      code: error.publicCode || 'SERVER_ERROR',
      message: error.publicCode ? error.message : '排行榜暫時無法使用，請聯絡管理者。'
    } };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
