// ============================================================
// ASTRO GURU TELEGRAM BOT — Complete Google Apps Script
// Paste this ENTIRE file into Extensions > Apps Script
// ============================================================

const TELEGRAM_API = (token) => `https://api.telegram.org/bot${token}`;
// ✅ Hardcode the /exec URL copied from "Manage deployments"
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzz_ImWDZkFLnAkrnYBJgGNIIGW2JsXqfXqPycNqKrGIUxEQUO0ZkzzfNmd9LRddjy0/exec'
// ─── CONFIG ─────────────────────────────────────────────────
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => { if (row[0]) config[row[0]] = row[1]; });
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
    if (!result.ok) {
      console.error(`Telegram API error [${method}]:`, JSON.stringify(result));
    }
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

function sendAudio(token, chatId, audioUrl, caption) {
  return callTelegram(token, 'sendAudio', {
    chat_id: chatId,
    audio: audioUrl,
    caption: caption || '',
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
        youtubeUrl: data[i][5],   // <- was audioUrl
        imageUrl: data[i][6]
      };
    }
  }
  return null;
}
/* Old code 
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
        audioUrl: data[i][5],
        imageUrl: data[i][6]
      };
    }
  }
  return null;
}
*/

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
        status: data[i][7]
      });
    }
  }
  return out;
}

// ─── WEBHOOK SETUP ──────────────────────────────────────────
function setWebhook() {
  const config = getConfig();
  const scriptUrl = ScriptApp.getService().getUrl();
  if (!scriptUrl || !scriptUrl.endsWith('/exec')) {
    console.error('❌ Deploy as Web App first. Current URL:', scriptUrl);
    return;
  }
  const result = callTelegram(config.BOT_TOKEN, 'setWebhook', {
    url: scriptUrl,
    allowed_updates: ['message', 'callback_query']
  });
  console.log('Webhook set:', JSON.stringify(result));
  return result;
}

function cleanupAndResetWebhook() {
  const config = getConfig();
  const token = config.BOT_TOKEN;

  console.log('=== CLEANUP & RESET WEBHOOK ===');

  const deleteResult = callTelegram(token, 'deleteWebhook', { drop_pending_updates: true });
  console.log('Delete successful:', deleteResult.ok);
  Utilities.sleep(2000);

  const currentUrl = WEBAPP_URL;
  console.log('Current URL:', currentUrl);

  if (!currentUrl || !currentUrl.endsWith('/exec')) {
    console.error('❌ URL invalid. Deploy as Web App first.');
    return;
  }

  const botCheck = callTelegram(token, 'getMe', {});
  if (!botCheck.ok) {
    console.error('❌ Invalid BOT_TOKEN');
    return;
  }
  console.log('✅ Bot:', '@' + botCheck.result.username);

  const setResult = callTelegram(token, 'setWebhook', {
    url: currentUrl,
    allowed_updates: ['message', 'callback_query']
  });
  if (!setResult.ok) {
    console.error('❌ Failed:', setResult.error_description);
    return;
  }
  console.log('✅ Webhook registered');

  Utilities.sleep(1000);
  const verify = callTelegram(token, 'getWebhookInfo', {});
  const info = verify.result;
  console.log('URL:', info.url);
  console.log('Pending updates:', info.pending_update_count);
  console.log('Last error:', info.last_error_message || 'none');

  if (info.url === currentUrl && !info.last_error_message) {
    console.log('✅✅✅ SUCCESS — send /start to @' + botCheck.result.username);
  }
}

function testCheckWebhook() {
  const config = getConfig();
  const result = callTelegram(config.BOT_TOKEN, 'getWebhookInfo', {});
  const info = result.result;
  console.log('=== WEBHOOK CHECK ===');
  console.log('URL:', info.url || 'EMPTY');
  console.log('Ends with /exec:', info.url ? info.url.endsWith('/exec') : false);
  console.log('Pending updates:', info.pending_update_count);
  console.log('Last error:', info.last_error_message || 'none');
}

function fullDiagnostic() {
  const config = getConfig();
  const token = config.BOT_TOKEN;
  console.log('=== DIAGNOSTIC START ===');
  console.log('BOT_TOKEN exists:', !!token);
  const botInfo = callTelegram(token, 'getMe', {});
  console.log('Bot reachable:', botInfo.ok);
  console.log('Bot username:', botInfo.result ? '@' + botInfo.result.username : 'FAILED');
  const wh = callTelegram(token, 'getWebhookInfo', {}).result;
  console.log('Webhook URL:', wh.url || 'EMPTY');
  console.log('Last error:', wh.last_error_message || 'none');
  const scriptUrl = ScriptApp.getService().getUrl();
  console.log('Script URL:', scriptUrl || 'EMPTY');
  console.log('=== DIAGNOSTIC END ===');
}

// ─── MAIN WEBHOOK HANDLER ───────────────────────────────────
function doPost(e) {
  const ok = ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const config = getConfig();
    const token = config.BOT_TOKEN;
    const update = JSON.parse(e.postData.contents);

    // Duplicate protection
    const cache = CacheService.getScriptCache();
    const key = 'upd_' + update.update_id;
    if (cache.get(key)) {
      console.log('Duplicate update ignored:', update.update_id);
      return ok;
    }
    cache.put(key, '1', 21600);

    console.log('Processing update:', update.update_id);

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
  const totalMessages = existingUser ? (Number(existingUser.data[12]) + 1) : 1;
  userInfo.totalMessages = totalMessages;

  logMessage(chatId, 'IN', text, 'received');

  if (text === '/start') {
    handleStart(token, chatId, userInfo, existingUser, config);
  } else if (text === '/status') {
    handleStatus(token, chatId, existingUser, config);
  } else if (text === '/today') {
    handleToday(token, chatId, existingUser, config);
  } else if (text === '/help') {
    handleHelp(token, chatId, config);
  } else if (text === '/schedule') {
    handleShowSchedule(token, chatId, config);
  } else {
    if (existingUser) {
      const updatedUser = buildUserObject(existingUser.data);
      updatedUser.totalMessages = totalMessages;
      updatedUser.lastMessageDate = new Date().toISOString();
      saveUser(updatedUser);
    }
    sendMessage(token, chatId, `I didn't understand that. Type /help for commands.`);
  }
}

// ─── /start ─────────────────────────────────────────────────
function handleStart(token, chatId, userInfo, existingUser, config) {
  const guruName = config.GURU_NAME || config.PLATFORM_NAME || 'Your Guru';
  const firstName = userInfo.firstName || 'there';

  // Save / update user
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
    // No schedule yet — guru hasn't assigned anything
    sendMessage(token, chatId,
      `🙏 *Namaste, ${firstName}!*\n\n` +
      `Welcome to *${guruName}*.\n\n` +
      `Your practice has not been set up yet. ` +
      `Please contact your guru to receive your mantra schedule.\n\n` +
      `Support: @${config.SUPPORT_USERNAME || 'admin'}`
    );
    logMessage(chatId, 'OUT', 'welcome — no schedule', 'sent');
    return;
  }

  // Schedule exists — show it, but don't let them pick
  const active = schedules.filter(s => String(s.status).toLowerCase() === 'active');
  const lines = active.map(s =>
    `• *${s.mantraKey}* — daily at ${s.time} IST`
  ).join('\n');

  sendMessage(token, chatId,
    `🙏 *Namaste, ${firstName}!*\n\n` +
    `Welcome back to *${guruName}*.\n\n` +
    `Your assigned practice:\n${lines}\n\n` +
    `I will send your mantra at the scheduled time. ` +
    `Use /today to receive it now.`,
    {
      reply_markup: {
        inline_keyboard: [
          [{ text: '📋 View my schedule', callback_data: 'show_schedule' }]
        ]
      }
    }
  );
  logMessage(chatId, 'OUT', 'welcome — schedule shown', 'sent');
}
/* Old code 
function handleStart(token, chatId, userInfo, existingUser, config) {
  const guruName = config.GURU_NAME || config.PLATFORM_NAME || 'Your Guru';
  const firstName = userInfo.firstName || 'there';

  const userData = {
    ...userInfo,
    paymentStatus: 'TRIAL',
    track: existingUser ? existingUser.data[5] : '',
    currentDay: existingUser ? existingUser.data[7] : 0,
    registrationDate: existingUser ? existingUser.data[8] : new Date().toISOString()
  };
  saveUser(userData);

  const welcome =
    `🙏 *Namaste, ${firstName}!*\n\n` +
    `Welcome to *${guruName}*.\n\n` +
    `I will send your daily mantra at your chosen time — ` +
    `with Sanskrit text, transliteration, meaning, and audio.\n\n` +
    `*Choose a mantra to begin:*`;

  sendMessage(token, chatId, welcome, {
    reply_markup: {
      inline_keyboard: [
        [{ text: '🕉️ Ganesha Mantra — daily', callback_data: 'mantra_ganesh' }],
        [{ text: '🌞 Gayatri Mantra — daily', callback_data: 'mantra_gayatri' }],
        [{ text: '💪 Hanuman Chalisa — Tue & Sat', callback_data: 'mantra_hanuman' }],
        [{ text: '📋 My current schedule', callback_data: 'show_schedule' }]
      ]
    }
  });

  logMessage(chatId, 'OUT', 'welcome + mantra menu', 'sent');
}*/

// ─── CALLBACK QUERY ─────────────────────────────────────────
function handleCallbackQuery(token, callbackQuery, config) {
  const chatId = callbackQuery.message.chat.id;
  const data = callbackQuery.data;

  callTelegram(token, 'answerCallbackQuery', {
    callback_query_id: callbackQuery.id
  });

  if (data === 'show_schedule') {
    handleShowSchedule(token, chatId, config);
  } else {
    console.warn('Unknown callback:', data);
  }
}
/*old
function handleCallbackQuery(token, callbackQuery, config) {
  const chatId = callbackQuery.message.chat.id;
  const data = callbackQuery.data;
  const from = callbackQuery.from || {};

  callTelegram(token, 'answerCallbackQuery', {
    callback_query_id: callbackQuery.id
  });

  if (data.startsWith('mantra_')) {
    handleMantraChosen(token, chatId, data.replace('mantra_', ''), from, config);
  } else if (data === 'show_schedule') {
    handleShowSchedule(token, chatId, config);
  } else if (data === 'change_schedule') {
    handleStart(token, chatId, {
      chatId,
      firstName: from.first_name || '',
      username: from.username || '',
      languageCode: from.language_code || '',
      lastMessageDate: new Date().toISOString(),
      totalMessages: 1
    }, getUserByChatId(chatId), config);
  } else {
    console.warn('Unknown callback:', data);
  }
}
*/
// ─── MANTRA CHOSEN ──────────────────────────────────────────
/*
function handleMantraChosen(token, chatId, mantraKey, from, config) {
  const mantra = getMantra(mantraKey);
  if (!mantra) {
    sendMessage(token, chatId, `⚠️ Mantra not found. Please try /start again.`);
    return;
  }

  const existingUser = getUserByChatId(chatId);
  const userData = existingUser ? buildUserObject(existingUser.data) : {
    chatId,
    firstName: from.first_name || '',
    lastName: from.last_name || '',
    username: from.username || '',
    languageCode: from.language_code || '',
    registrationDate: new Date().toISOString()
  };
  userData.track = mantraKey;
  userData.paymentStatus = 'TRIAL';
  userData.currentDay = Number(userData.currentDay) || 0;
  userData.lastMessageDate = new Date().toISOString();
  saveUser(userData);

  // Default schedule: 06:30 daily for 30 days
  const start = new Date();
  const end = new Date(); end.setDate(end.getDate() + 30);
  const fmt = d => Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
  const scheduleId = 'S' + Date.now();
  getSheet('Schedules').appendRow([
    scheduleId, chatId, mantraKey, '06:30', 'daily',
    fmt(start), fmt(end), 'active'
  ]);

  sendMessage(token, chatId,
    `✅ *${mantra.name}* — added to your daily practice.\n\n` +
    `⏰ Every day at *6:30 AM IST*\n` +
    `📅 For the next 30 days\n\n` +
    `Here is your first recitation 👇`
  );

  sendMantraMessage(token, chatId, mantra, config);
  logMessage(chatId, 'OUT', `mantra scheduled: ${mantraKey}`, 'sent');
}
*/

// Run this manually from the Apps Script editor to assign a mantra to a client
function assignMantraToClient(clientChatId, mantraKey, timeIst, days, durationDays) {
  const fmt = d => Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
  const start = new Date();
  const end = new Date(); end.setDate(end.getDate() + (durationDays || 30));
  const scheduleId = 'S' + Date.now();
  getSheet('Schedules').appendRow([
    scheduleId, clientChatId, mantraKey, timeIst || '06:30',
    days || 'daily', fmt(start), fmt(end), 'active'
  ]);
  console.log('Assigned', mantraKey, 'to', clientChatId, 'at', timeIst);
}
// ─── SEND MANTRA ────────────────────────────────────────────

function sendMantraMessage(token, chatId, mantra, config) {
  const text =
    `🕉️ *${mantra.name}*\n\n` +
    `*Sanskrit:*\n${mantra.script}\n\n` +
    `*Transliteration:*\n_${mantra.transliteration}_\n\n` +
    `*Meaning:*\n${mantra.meaning}\n\n` +
    `Chant with devotion 🙏`;

  // Build YouTube button if present
  const buttons = [];
  if (mantra.youtubeUrl && String(mantra.youtubeUrl).startsWith('http')) {
    buttons.push([{ text: '▶️ Listen on YouTube', url: mantra.youtubeUrl }]);
  }
  const replyMarkup = buttons.length
    ? { reply_markup: { inline_keyboard: buttons } }
    : {};

  if (mantra.imageUrl && String(mantra.imageUrl).startsWith('http')) {
    // Photo caption max 1024 chars — truncate if needed
    const caption = text.length > 1024 ? text.substring(0, 1020) + '…' : text;
    sendPhoto(token, chatId, mantra.imageUrl, caption);
    if (buttons.length) {
      sendMessage(token, chatId, '👇', { reply_markup: { inline_keyboard: buttons } });
    }
  } else {
    sendMessage(token, chatId, text, replyMarkup);
  }

  // increment count
  const user = getUserByChatId(chatId);
  if (user) {
    const d = buildUserObject(user.data);
    d.currentDay = Number(d.currentDay || 0) + 1;
    saveUser(d);
  }
}
/* Old code
function sendMantraMessage(token, chatId, mantra, config) {
  const text =
    `🕉️ *${mantra.name}*\n\n` +
    `*Sanskrit:*\n${mantra.script}\n\n` +
    `*Transliteration:*\n_${mantra.transliteration}_\n\n` +
    `*Meaning:*\n${mantra.meaning}\n\n` +
    `Chant with devotion 🙏`;

  if (mantra.imageUrl && String(mantra.imageUrl).startsWith('http')) {
    sendPhoto(token, chatId, mantra.imageUrl, text.substring(0, 1024));
  } else {
    sendMessage(token, chatId, text);
  }

  if (mantra.audioUrl && String(mantra.audioUrl).startsWith('http')) {
    sendAudio(token, chatId, mantra.audioUrl, mantra.name);
  }

  const user = getUserByChatId(chatId);
  if (user) {
    const d = buildUserObject(user.data);
    d.currentDay = Number(d.currentDay || 0) + 1;
    saveUser(d);
  }
}
*/

// ─── /status ────────────────────────────────────────────────

function handleStatus(token, chatId, existingUser, config) {
  if (!existingUser) {
    sendMessage(token, chatId, `You're not registered. Send /start to begin.`);
    return;
  }
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
/* Old 
function handleStatus(token, chatId, existingUser, config) {
  if (!existingUser) {
    sendMessage(token, chatId, `You're not registered. Send /start to begin.`);
    return;
  }
  const d = existingUser.data;
  const schedules = getSchedulesForUser(chatId);
  const lines = schedules.length
    ? schedules.map(s => `• ${s.mantraKey} at ${s.time} IST — ${s.days}`).join('\n')
    : '• No active schedule';

  sendMessage(token, chatId,
    `📊 *Your practice*\n\n` +
    `👤 ${d[1]} ${d[2]}\n` +
    `🕉️ Total recitations sent: ${d[7] || 0}\n\n` +
    `*Active schedules:*\n${lines}\n\n` +
    `Use /start to add another mantra.`
  );
}
*/
// ─── /today ─────────────────────────────────────────────────
function handleToday(token, chatId, existingUser, config) {
  if (!existingUser || !existingUser.data[5]) {
    sendMessage(token, chatId, `No mantra selected yet. Send /start.`);
    return;
  }
  const mantra = getMantra(existingUser.data[5]);
  if (!mantra) {
    sendMessage(token, chatId, `⚠️ Mantra not found. Send /start to reset.`);
    return;
  }
  sendMantraMessage(token, chatId, mantra, config);
}

// ─── /help ──────────────────────────────────────────────────
function handleHelp(token, chatId, config) {
  sendMessage(token, chatId,
    `🤖 *Commands*\n\n` +
    `/start — choose a mantra\n` +
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
/* Old code 
function handleShowSchedule(token, chatId, config) {
  const schedules = getSchedulesForUser(chatId);
  if (!schedules.length) {
    sendMessage(token, chatId, `No active schedules. Send /start to add a mantra.`);
    return;
  }
  const lines = schedules.map(s => {
    let endStr = '';
    try {
      const end = s.endDate instanceof Date ? s.endDate : new Date(s.endDate);
      endStr = Utilities.formatDate(end, 'Asia/Kolkata', 'dd MMM yyyy');
    } catch (e) { endStr = String(s.endDate); }
    return `• ${s.mantraKey} — ${s.time} IST (${s.days}) until ${endStr}`;
  }).join('\n');

  sendMessage(token, chatId,
    `📅 *Your schedules:*\n\n${lines}\n\nTo change: /start`,
    {
      reply_markup: {
        inline_keyboard: [[{ text: '➕ Add another mantra', callback_data: 'change_schedule' }]]
      }
    }
  );
}*/

// ─── DAILY SCHEDULER ────────────────────────────────────────
function sendDailyMantras() {
  const config = getConfig();
  const token = config.BOT_TOKEN;
  const now = new Date();
  const istNow = new Date(now.getTime() + (5.5 * 60 * 60 * 1000));
  const hhmm = Utilities.formatDate(istNow, 'GMT', 'HH:mm');
  const todayAbbr = Utilities.formatDate(istNow, 'GMT', 'EEE').toLowerCase().substring(0, 3);
  const todayDate = Utilities.formatDate(istNow, 'GMT', 'yyyy-MM-dd');

  const sheet = getSheet('Schedules');
  const data = sheet.getDataRange().getValues();
  let sent = 0, skipped = 0, errors = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[7]).toLowerCase();
    const time = String(row[3]);
    const days = String(row[4]).toLowerCase();
    const startDate = row[5] instanceof Date
      ? Utilities.formatDate(row[5], 'GMT', 'yyyy-MM-dd')
      : String(row[5]);
    const endDate = row[6] instanceof Date
      ? Utilities.formatDate(row[6], 'GMT', 'yyyy-MM-dd')
      : String(row[6]);

    if (status !== 'active') {  
        console.log(`📊 Status — sent: ${status} `);
         skipped++; continue; }
   // if (time !== hhmm) { console.log(`📊 time — sent: ${time} `); skipped++; continue; }
    if (todayDate < startDate || todayDate > endDate) { skipped++; continue; }

    const dayMatch = (days === 'daily') || days.split(',').map(s => s.trim()).includes(todayAbbr);
    if (!dayMatch) { skipped++; continue; }

    const mantra = getMantra(row[2]);
    if (!mantra) { errors++; continue; }

    try {
      sendMantraMessage(token, row[1], mantra, config);
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

  ScriptApp.newTrigger('sendDailyMantras')
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log('✅ Scheduler trigger installed — runs every 5 min');
}

// ─── TEST FUNCTIONS ─────────────────────────────────────────
function getRecentUpdates() {
  const config = getConfig();
  callTelegram(config.BOT_TOKEN, 'deleteWebhook', {});
  const url = `${TELEGRAM_API(config.BOT_TOKEN)}/getUpdates`;
  const response = UrlFetchApp.fetch(url);
  const result = JSON.parse(response.getContentText());
  console.log('Recent updates:');
  result.result.forEach(u => {
    if (u.message) {
      console.log(`chatId: ${u.message.chat.id} | name: ${u.message.from.first_name} | text: ${u.message.text}`);
    }
  });
  console.log('⚠️ After copying chatId, run cleanupAndResetWebhook() to restore webhook');
}

function testMessageToSelf() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config sheet'); return; }
  sendMessage(config.BOT_TOKEN, chatId,
    '✅ *Bot is working!*\n\nWebhook and sending are live.\n\nTry /start in the bot.'
  );
  console.log('Test message sent to', chatId);
}

function testMantraMenu() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config'); return; }
  handleStart(config.BOT_TOKEN, chatId,
    { chatId, firstName: 'Test', lastName: '', username: 'test', languageCode: 'en' },
    getUserByChatId(chatId), config);
  console.log('Sent mantra menu');
}

function testSendGaneshNow() {
  const config = getConfig();
  const chatId = config.TEST_CHAT_ID;
  if (!chatId) { console.log('Add TEST_CHAT_ID to Config'); return; }
  const mantra = getMantra('ganesh');
  if (!mantra) { console.log('Mantra "ganesh" not found in Mantras sheet'); return; }
  sendMantraMessage(config.BOT_TOKEN, chatId, mantra, config);
  console.log('Sent Ganesh mantra');
}

function testRunSchedulerNow() {
  console.log('Running scheduler manually...');
  sendDailyMantras();
}
function forceSetWebhookWithExecUrl() {
  const config = getConfig();
  const token = config.BOT_TOKEN;
  
  // ⚠️ Manually replace this with the /exec URL you copied from "Manage deployments"
  const execUrl = 'https://script.google.com/macros/s/AKfycbzz_ImWDZkFLnAkrnYBJgGNIIGW2JsXqfXqPycNqKrGIUxEQUO0ZkzzfNmd9LRddjy0/exec';
  
  // First delete the old webhook
  callTelegram(token, 'deleteWebhook', { drop_pending_updates: true });
  Utilities.sleep(2000);
  
  // Re-register with /exec
  const result = callTelegram(token, 'setWebhook', {
    url: execUrl,
    allowed_updates: ['message', 'callback_query']
  });
  
  console.log('Set webhook result:', JSON.stringify(result));
  
  // Verify
  Utilities.sleep(1000);
  const info = callTelegram(token, 'getWebhookInfo', {});
  console.log('Webhook URL:', info.result.url);
  console.log('Last error:', info.result.last_error_message || 'none');
}

function forceSetWebhook() {
  const config  = getConfig();
  const token   = config.BOT_TOKEN;

  // Get the current script's URL
  const url = ScriptApp.getService().getUrl();
  console.log('Registering URL:', url);

  if (!url || url === '') {
    console.error('❌ URL is empty — make sure you deployed as a Web App first');
    return;
  }

  // Delete any existing webhook first
  callTelegram(token, 'deleteWebhook', { drop_pending_updates: true });
  Utilities.sleep(1000);

  // Set the new webhook
  const result = callTelegram(token, 'setWebhook', {
    url: url,
    allowed_updates: ['message', 'callback_query']
  });

  console.log('Set webhook result:', JSON.stringify(result));
  // Must see: {"ok":true,"result":true,"description":"Webhook was set"}
}

function testListSchedules() {
  const sheet = getSheet('Schedules');
  const data = sheet.getDataRange().getValues();
  console.log('=== SCHEDULES ===');
  for (let i = 1; i < data.length; i++) {
    console.log(`Row ${i + 1}:`, data[i].join(' | '));
  }
}

function testCheckBot() {
  const config = getConfig();
  const result = callTelegram(config.BOT_TOKEN, 'getMe', {});
  console.log('Bot info:', JSON.stringify(result.result, null, 2));
}
