// ============================================================
// ASTRO GURU TELEGRAM BOT — Complete Google Apps Script
// Paste this ENTIRE file into Extensions > Apps Script
// ============================================================

const TELEGRAM_API = (token) => `https://api.telegram.org/bot${token}`;
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzz_ImWDZkFLnAkrnYBJgGNIIGW2JsXqfXqPycNqKrGIUxEQUO0ZkzzfNmd9LRddjy0/exec';

// ─── CONFIG ─────────────────────────────────────────────────
let _configCache = null;
function getConfig(force) {
  if (_configCache && !force) return _configCache;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
  _configCache = config;
  return config;
}

// ─── SHEET HELPERS ──────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getUserByChatId(chatId) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(chatId)) {
      return { rowIndex: i + 1, data: data[i] };
    }
  }
  return null;
}

function saveUser(user) {
  const sheet = getSheet('Users');
  const existing = getUserByChatId(user.chatId);
  const row = [
    user.chatId,
    user.firstName || '',
    user.lastName || '',
    user.username || '',
    user.languageCode || '',
    user.track || '',
    user.paymentStatus || 'TRIAL',
    user.currentDay || 0,
    user.registrationDate || new Date().toISOString(),
    user.lastMessageDate || new Date().toISOString(),
    user.source || 'organic',
    user.deviceInfo || '',
    user.totalMessages || 1,
    user.notes || ''
  ];
  if (existing) {
    sheet.getRange(existing.rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function buildUserObject(row) {
  return {
    chatId: row[0],
    firstName: row[1],
    lastName: row[2],
    username: row[3],
    languageCode: row[4],
    track: row[5],
    paymentStatus: row[6],
    currentDay: row[7],
    registrationDate: row[8],
    lastMessageDate: row[9],
    source: row[10],
    deviceInfo: row[11],
    totalMessages: row[12],
    notes: row[13]
  };
}

function logMessage(chatId, direction, messageText, status) {
  try {
    const sheet = getSheet('MessageLog');
    if (!sheet) return;
    sheet.appendRow([
      new Date().toISOString(),
      chatId,
      direction,
      String(messageText || '').substring(0, 500),
      status
    ]);
  } catch (e) {
    console.error('logMessage failed:', e.message);
  }
}

// ─── TELEGRAM API ───────────────────────────────────────────
function callTelegram(token, method, payload) {
  const url = `${TELEGRAM_API(token)}/${method}`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) console.error(`Telegram API error [${method}]:`, JSON.stringify(result));
    return result;
  } catch (e) {
    console.error(`Network error [${method}]:`, e.message);
    return { ok: false, error: e.message };
  }
}

function sendMessage(token, chatId, text, extra) {
  return callTelegram(token, 'sendMessage', {
    chat_id: chatId,
    text: text,
    parse_mode: 'Markdown',
    ...extra
  });
}

function sendPhoto(token, chatId, photoUrl, caption) {
  return callTelegram(token, 'sendPhoto', {
    chat_id: chatId,
    photo: photoUrl,
    caption: caption,
    parse_mode: 'Markdown',
    protect_content: true
  });
}

// ─── MANTRA / SCHEDULE LOOKUPS ──────────────────────────────
function getMantra(mantraKey) {
  const sheet = getSheet('Mantras');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(mantraKey)) {
      return {
        key: data[i][0],
        name: data[i][1],
        script: data[i][2],
        transliteration: data[i][3],
        meaning: data[i][4],
        youtubeUrl: data[i][5],
        imageUrl: data[i][6]
      };
    }
  }
  return null;
}

function getSchedulesForUser(chatId) {
  const sheet = getSheet('Schedules');
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(chatId)) {
      out.push({
        rowIndex: i + 1,
        scheduleId: data[i][0],
        chatId: data[i][1],
        mantraKey: data[i][2],
        time: data[i][3],
        days: data[i][4],
        startDate: data[i][5],
        endDate: data[i][6],
        status: data[i][7],
        lastSentDate: data[i][8] || ''
      });
    }
  }
  return out;
}

// ─── WEBHOOK ────────────────────────────────────────────────
function cleanupAndResetWebhook() {
  const config = getConfig(true);
  const token = config.BOT_TOKEN;

  console.log('=== CLEANUP & RESET WEBHOOK ===');
  callTelegram(token, 'deleteWebhook', { drop_pending_updates: true });
  Utilities.sleep(2000);

  if (!WEBAPP_URL || !WEBAPP_URL.endsWith('/exec')) {
    console.error('❌ WEBAPP_URL invalid:', WEBAPP_URL);
    return;
  }

  const botCheck = callTelegram(token, 'getMe', {});
  if (!botCheck.ok) { console.error('❌ Invalid BOT_TOKEN'); return; }
  console.log('✅ Bot:', '@' + botCheck.result.username);

  const setResult = callTelegram(token, 'setWebhook', {
    url: WEBAPP_URL,
    allowed_updates: ['message', 'callback_query']
  });
  if (!setResult.ok) { console.error('❌ Failed:', setResult.error_description); return; }
  console.log('✅ Webhook registered');

  Utilities.sleep(1000);
  const info = callTelegram(token, 'getWebhookInfo', {}).result;
  console.log('URL:', info.url);
  console.log('Pending updates:', info.pending_update_count);
  console.log('Last error:', info.last_error_message || 'none');
  if (info.url === WEBAPP_URL && !info.last_error_message) {
    console.log('✅✅✅ SUCCESS — send /start to @' + botCheck.result.username);
  }
}

function testCheckWebhook() {
  const config = getConfig(true);
  const info = callTelegram(config.BOT_TOKEN, 'getWebhookInfo', {}).result;
  console.log('=== WEBHOOK CHECK ===');
  console.log('URL:', info.url || 'EMPTY');
  console.log('Pending updates:', info.pending_update_count);
  console.log('Last error:', info.last_error_message || 'none');
}

// ─── MAIN WEBHOOK HANDLER ───────────────────────────────────
function doPost(e) {
  const ok = ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const config = getConfig(true);
    const token = config.BOT_TOKEN;
    const update = JSON.parse(e.postData.contents);

    const cache = CacheService.getScriptCache();
    const key = 'upd_' + update.update_id;
    if (cache.get(key)) { console.log('Duplicate update ignored:', update.update_id); return ok; }
    cache.put(key, '1', 21600);

    if (update.message) handleMessage(token, update.message, config);
    if (update.callback_query) handleCallbackQuery(token, update.callback_query, config);
  } catch (err) {
    console.error('doPost ERROR:', err.message, err.stack);
  }
  return ok;
}

// ─── MESSAGE HANDLER ────────────────────────────────────────
function handleMessage(token, message, config) {
  const chatId = message.chat.id;
  const text = message.text || '';
  const from = message.from || {};

  const userInfo = {
    chatId: chatId,
    firstName: from.first_name || '',
    lastName: from.last_name || '',
    username: from.username || '',
    languageCode: from.language_code || '',
    deviceInfo: message.via_bot ? 'bot' : 'direct',
    lastMessageDate: new Date().toISOString(),
    source: 'organic'
  };

  const existingUser = getUserByChatId(chatId);
  userInfo.totalMessages = existingUser ? (Number(existingUser.data[12]) + 1) : 1;
  logMessage(chatId, 'IN', text, 'received');

  if (text === '/start')        handleStart(token, chatId, userInfo, existingUser, config);
  else if (text === '/status')  handleStatus(token, chatId, existingUser, config);
  else if (text === '/today')   handleToday(token, chatId, existingUser, config);
  else if (text === '/help')    handleHelp(token, chatId, config);
  else if (text === '/schedule') handleShowSchedule(token, chatId, config);
  else {
    if (existingUser) {
      const u = buildUserObject(existingUser.data);
      u.totalMessages = userInfo.totalMessages;
      u.lastMessageDate = new Date().toISOString();
      saveUser(u);
    }
    sendMessage(token, chatId, `I didn't understand that. Type /help for commands.`);
  }
}

// ─── /start ─────────────────────────────────────────────────
function handleStart(token, chatId, userInfo, existingUser, config) {
  const guruName = config.GURU_NAME || config.PLATFORM_NAME || 'Your Guru';
  const firstName = userInfo.firstName || 'there';

  const userData = {
    ...userInfo,
    paymentStatus: existingUser ? existingUser.data[6] : 'TRIAL',
    track: existingUser ? existingUser.data[5] : '',
    currentDay: existingUser ? existingUser.data[7] : 0,
    registrationDate: existingUser ? existingUser.data[8] : new Date().toISOString()
  };
  saveUser(userData);

  const schedules = getSchedulesForUser(chatId);

  if (!schedules.length) {
    sendMessage(token, chatId,
      `🙏 *Namaste, ${firstName}!*\n\n` +
      `Welcome to *${guruName}*.\n\n` +
      `Your practice has not been set up yet. Please contact your guru to receive your mantra schedule.\n\n` +
      `Support: @${config.SUPPORT_USERNAME || 'admin'}`
    );
    logMessage(chatId, 'OUT', 'welcome — no schedule', 'sent');
    return;
  }

  const active = schedules.filter(s => String(s.status).toLowerCase() === 'active');
  const lines = active.map(s => `• *${s.mantraKey}* — daily at ${s.time} IST`).join('\n');

  sendMessage(token, chatId,
    `🙏 *Namaste, ${firstName}!*\n\n` +
    `Welcome back to *${guruName}*.\n\n` +
    `Your assigned practice:\n${lines}\n\n` +
    `I will send your mantra at the scheduled time. Use /today to receive it now.`,
    { reply_markup: { inline_keyboard: [[{ text: '📋 View my schedule', callback_data: 'show_schedule' }]] } }
  );
  logMessage(chatId, 'OUT', 'welcome — schedule shown', 'sent');
}

// ─── CALLBACK QUERY ─────────────────────────────────────────
function handleCallbackQuery(token, callbackQuery, config) {
  const chatId = callbackQuery.message.chat.id;
  const data = callbackQuery.data;
  callTelegram(token, 'answerCallbackQuery', { callback_query_id: callbackQuery.id });

  if (data === 'show_schedule') handleShowSchedule(token, chatId, config);
  else console.warn('Unknown callback:', data);
}

// ─── SEND MANTRA ────────────────────────────────────────────
function sendMantraMessage(token, chatId, mantra, config) {
  const text =
    `🕉️ *${mantra.name}*\n\n` +
    `*Sanskrit:*\n${mantra.script}\n\n` +
    `*Transliteration:*\n_${mantra.transliteration}_\n\n` +
    `*Meaning:*\n${mantra.meaning}\n\n` +
    `Chant with devotion 🙏`;

  const buttons = [];
  if (mantra.youtubeUrl && String(mantra.youtubeUrl).startsWith('http')) {
    buttons.push([{ text: '▶️ Listen on YouTube', url: mantra.youtubeUrl }]);
  }

  if (mantra.imageUrl && String(mantra.imageUrl).startsWith('http')) {
    const caption = text.length > 1024 ? text.substring(0, 1020) + '…' : text;
    sendPhoto(token, chatId, mantra.imageUrl, caption);
    if (buttons.length) {
      sendMessage(token, chatId, '👇', { reply_markup: { inline_keyboard: buttons } });
    }
  } else {
    const extra = buttons.length ? { reply_markup: { inline_keyboard: buttons } } : {};
    sendMessage(token, chatId, text, extra);
  }

  const user = getUserByChatId(chatId);
  if (user) {
    const d = buildUserObject(user.data);
    d.currentDay = Number(d.currentDay || 0) + 1;
    saveUser(d);
  }
}

// ─── /status ────────────────────────────────────────────────
function handleStatus(token, chatId, existingUser, config) {
  if (!existingUser) { sendMessage(token, chatId, `You're not registered. Send /start to begin.`); return; }
  const d = existingUser.data;
  const schedules = getSchedulesForUser(chatId).filter(s => String(s.status).toLowerCase() === 'active');

  const lines = schedules.length
    ? schedules.map(s => `• *${s.mantraKey}* — ${s.time} IST (${s.days})`).join('\n')
    : '• No active practice assigned yet. Contact your guru.';

  sendMessage(token, chatId,
    `📊 *Your practice*\n\n` +
    `👤 ${d[1]} ${d[2] || ''}\n` +
    `🕉️ Recitations received: ${d[7] || 0}\n\n` +
    `*Assigned mantras:*\n${lines}\n\n` +
    `Support: @${config.SUPPORT_USERNAME || 'admin'}`
  );
}

// ─── /today ─────────────────────────────────────────────────
function handleToday(token, chatId, existingUser, config) {
  if (!existingUser || !existingUser.data[5]) {
    sendMessage(token, chatId, `No mantra selected yet. Send /start.`);
    return;
  }
  const mantra = getMantra(existingUser.data[5]);
  if (!mantra) { sendMessage(token, chatId, `⚠️ Mantra not found. Send /start to reset.`); return; }
  sendMantraMessage(token, chatId, mantra, config);
}

// ─── /help ──────────────────────────────────────────────────
function handleHelp(token, chatId, config) {
  sendMessage(token, chatId,
    `🤖 *Commands*\n\n` +
    `/start — begin or view your practice\n` +
    `/today — receive today's mantra now\n` +
    `/status — see your practice\n` +
    `/schedule — see active schedules\n` +
    `/help — this menu\n\n` +
    `Support: @${config.SUPPORT_USERNAME || 'admin'}`
  );
}

// ─── /schedule ──────────────────────────────────────────────
function handleShowSchedule(token, chatId, config) {
  const schedules = getSchedulesForUser(chatId).filter(s => String(s.status).toLowerCase() === 'active');
  if (!schedules.length) {
    sendMessage(token, chatId,
      `No active practice assigned.\n\nPlease contact your guru: @${config.SUPPORT_USERNAME || 'admin'}`
    );
    return;
  }
  const lines = schedules.map(s => {
    let endStr = '';
    try {
      const end = s.endDate instanceof Date ? s.endDate : new Date(s.endDate);
      endStr = Utilities.formatDate(end, 'Asia/Kolkata', 'dd MMM yyyy');
    } catch (e) { endStr = String(s.endDate); }
    return `• *${s.mantraKey}* — ${s.time} IST (${s.days}) until ${endStr}`;
  }).join('\n');

  sendMessage(token, chatId,
    `📅 *Your schedule:*\n\n${lines}\n\n` +
    `To change your practice, contact your guru: @${config.SUPPORT_USERNAME || 'admin'}`
  );
}

// ─── DAILY SCHEDULER (production) ───────────────────────────
function sendDailyMantras() {
  const config = getConfig(true);
  const token = config.BOT_TOKEN;
  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const hhmm = Utilities.formatDate(istNow, 'GMT', 'HH:mm');
  const todayAbbr = Utilities.formatDate(istNow, 'GMT', 'EEE').toLowerCase().substring(0, 3);
  const todayDate = Utilities.formatDate(istNow, 'GMT', 'yyyy-MM-dd');
  const nowMin = parseInt(hhmm.split(':')[0], 10) * 60 + parseInt(hhmm.split(':')[1], 10);

  const sheet = getSheet('Schedules');
  const data = sheet.getDataRange().getValues();
  let sent = 0, skipped = 0, errors = 0;

  console.log(`📤 Daily send @ ${hhmm} IST (${todayAbbr})`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[7] || '').toLowerCase();
    const time   = String(row[3] || '').trim();
    const days   = String(row[4] || '').toLowerCase().trim();
    const lastSent = String(row[8] || '');
    const startDate = row[5] instanceof Date
      ? Utilities.formatDate(row[5], 'GMT', 'yyyy-MM-dd')
      : String(row[5] || '');
    const endDate = row[6] instanceof Date
      ? Utilities.formatDate(row[6], 'GMT', 'yyyy-MM-dd')
      : String(row[6] || '');

    if (status !== 'active')                    { skipped++; continue; }
    if (lastSent === todayDate)                 { skipped++; continue; }
    if (!/^\d{1,2}:\d{2}$/.test(time))          { skipped++; continue; }
    const schMin = parseInt(time.split(':')[0], 10) * 60 + parseInt(time.split(':')[1], 10);
    if (nowMin < schMin || nowMin - schMin > 15){ skipped++; continue; }
    if (todayDate < startDate || todayDate > endDate) { skipped++; continue; }

    const dayMatch = (days === 'daily') || days.split(',').map(s => s.trim()).includes(todayAbbr);
    if (!dayMatch) { skipped++; continue; }

    const mantra = getMantra(row[2]);
    if (!mantra) { errors++; continue; }

    try {
      sendMantraMessage(token, row[1], mantra, config);
      sheet.getRange(i + 1, 9).setValue(todayDate);   // ← inside the loop, per-row
      sent++;
      Utilities.sleep(150);
    } catch (e) {
      console.error('sendDailyMantras error for', row[1], e.message);
      errors++;
    }
  }

  console.log(`📊 Mantra send — sent: ${sent} | skipped: ${skipped} | errors: ${errors}`);
  logMessage('SYSTEM', 'OUT', `Mantra send: ${sent} sent`, 'done');
}

function installSchedulerTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyMantras')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyMantras').timeBased().everyMinutes(5).create();
  console.log('✅ Scheduler trigger installed — every 5 min');
}

// ─── TEST SCHEDULER (sends only to TEST_CHAT_ID) ────────────
function testSchedulerTick() {
  const config = getConfig(true);
  const token = config.BOT_TOKEN;
  const testChatId = config.TEST_CHAT_ID;
  if (!testChatId) { console.log('⚠️ TEST_CHAT_ID missing'); return; }

  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const hhmm = Utilities.formatDate(istNow, 'GMT', 'HH:mm');
  const todayAbbr = Utilities.formatDate(istNow, 'GMT', 'EEE').toLowerCase().substring(0, 3);
  const todayDate = Utilities.formatDate(istNow, 'GMT', 'yyyy-MM-dd');
  const nowMin = parseInt(hhmm.split(':')[0], 10) * 60 + parseInt(hhmm.split(':')[1], 10);

  console.log(`🧪 TEST TICK @ ${hhmm} IST (${todayAbbr}, ${todayDate})`);

  const data = getSheet('Schedules').getDataRange().getValues();
  let matched = 0, sent = 0, skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[7] || '').toLowerCase();
    const time   = String(row[3] || '').trim();
    const days   = String(row[4] || '').toLowerCase().trim();
    const start  = row[5] instanceof Date ? Utilities.formatDate(row[5], 'GMT', 'yyyy-MM-dd') : String(row[5] || '');
    const end    = row[6] instanceof Date ? Utilities.formatDate(row[6], 'GMT', 'yyyy-MM-dd') : String(row[6] || '');

    if (status !== 'active')                    { skipped++; continue; }
    if (!/^\d{1,2}:\d{2}$/.test(time))          { skipped++; continue; }
    const schMin = parseInt(time.split(':')[0], 10) * 60 + parseInt(time.split(':')[1], 10);
    if (nowMin < schMin || nowMin - schMin > 15){ skipped++; continue; }
    if (todayDate < start || todayDate > end)   { skipped++; continue; }

    const dayMatch = (days === 'daily') || days.split(',').map(s => s.trim()).includes(todayAbbr);
    if (!dayMatch) { skipped++; continue; }

    const mantra = getMantra(row[2]);
    if (!mantra) { skipped++; continue; }

    matched++;
    try {
      sendMantraMessage(token, testChatId, mantra, config);
      sent++;
      console.log(`✅ Test-sent "${row[2]}" (would go to ${row[1]})`);
    } catch (e) {
      console.error('testSchedulerTick send error:', e.message);
    }
  }

  console.log(`🧪 TEST TICK done — matched: ${matched} | sent-to-test: ${sent} | skipped: ${skipped}`);
}

function installTestSchedulerTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'testSchedulerTick')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('testSchedulerTick').timeBased().everyMinutes(5).create();
  console.log('✅ Test trigger installed — every 5 min → TEST_CHAT_ID only');
}

function uninstallTestSchedulerTrigger() {
  const n = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'testSchedulerTick')
    .reduce((acc, t) => { ScriptApp.deleteTrigger(t); return acc + 1; }, 0);
  console.log(`🗑 Removed ${n} test trigger(s)`);
}

function listAllTriggers() {
  const ts = ScriptApp.getProjectTriggers();
  if (!ts.length) { console.log('No triggers installed'); return; }
  ts.forEach(t => console.log(`• ${t.getHandlerFunction()} — ${t.getEventType()}`));
}

// ─── TEST / DIAGNOSTIC HELPERS ──────────────────────────────
function getRecentUpdates() {
  const config = getConfig(true);
  callTelegram(config.BOT_TOKEN, 'deleteWebhook', {});
  const response = UrlFetchApp.fetch(`${TELEGRAM_API(config.BOT_TOKEN)}/getUpdates`);
  const result = JSON.parse(response.getContentText());
  result.result.forEach(u => {
    if (u.message) console.log(`chatId: ${u.message.chat.id} | ${u.message.from.first_name} | ${u.message.text}`);
  });
  cleanupAndResetWebhook();   // ← always restore
}

function testMessageToSelf() {
  const config = getConfig(true);
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID'); return; }
  sendMessage(config.BOT_TOKEN, chatId, '✅ *Bot is working!*\n\nTry /start.');
  console.log('Sent to', chatId);
}

function testCheckBot() {
  const config = getConfig(true);
  const r = callTelegram(config.BOT_TOKEN, 'getMe', {});
  console.log('Bot:', JSON.stringify(r.result, null, 2));
}

function testListSchedules() {
  const data = getSheet('Schedules').getDataRange().getValues();
  console.log('=== SCHEDULES ===');
  for (let i = 1; i < data.length; i++) console.log(`Row ${i + 1}:`, data[i].join(' | '));
}

function testRunSchedulerNow() {
  console.log('Running production scheduler manually...');
  sendDailyMantras();
}

function diagnoseSchedules() {
  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const hhmm = Utilities.formatDate(istNow, 'GMT', 'HH:mm');
  const todayAbbr = Utilities.formatDate(istNow, 'GMT', 'EEE').toLowerCase().substring(0, 3);
  const todayDate = Utilities.formatDate(istNow, 'GMT', 'yyyy-MM-dd');
  const nowMin = parseInt(hhmm.split(':')[0], 10) * 60 + parseInt(hhmm.split(':')[1], 10);

  console.log('NOW (IST):', hhmm, '| day:', todayAbbr, '| date:', todayDate);

  const data = getSheet('Schedules').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const status = String(r[7] || '').toLowerCase();
    const time   = String(r[3] || '').trim();
    const days   = String(r[4] || '').toLowerCase().trim();
    const start  = r[5] instanceof Date ? Utilities.formatDate(r[5], 'GMT', 'yyyy-MM-dd') : String(r[5] || '');
    const end    = r[6] instanceof Date ? Utilities.formatDate(r[6], 'GMT', 'yyyy-MM-dd') : String(r[6] || '');
    const lastSent = String(r[8] || '');

    const reasons = [];
    if (status !== 'active') reasons.push(`status="${r[7]}"`);
    if (lastSent === todayDate) reasons.push('already sent today');
    if (!/^\d{1,2}:\d{2}$/.test(time)) reasons.push(`time="${time}" malformed`);
    else {
      const schMin = parseInt(time.split(':')[0], 10) * 60 + parseInt(time.split(':')[1], 10);
      if (nowMin < schMin || nowMin - schMin > 15) reasons.push(`time "${time}" not in ±15min window (now ${hhmm})`);
    }
    if (todayDate < start) reasons.push(`start ${start} in future`);
    if (todayDate > end)   reasons.push(`end ${end} passed`);
    const dayMatch = (days === 'daily') || days.split(',').map(s => s.trim()).includes(todayAbbr);
    if (!dayMatch) reasons.push(`days="${days}" lacks "${todayAbbr}"`);

    console.log(`Row ${i + 1} | chatId=${r[1]} | mantra=${r[2]} | ` + (reasons.length ? '❌ ' + reasons.join(' · ') : '✅ would fire'));
  }
}

function testSendGaneshNow() {
  const config = getConfig(true);
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID'); return; }
  const m = getMantra('ganesh');
  if (!m) { console.log('Mantra "ganesh" not found'); return; }
  sendMantraMessage(config.BOT_TOKEN, chatId, m, config);
  console.log('Sent Ganesh mantra');
}
