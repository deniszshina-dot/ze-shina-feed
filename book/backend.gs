// ZE-Shina: запис на шиномонтаж + пошук зберігання + Telegram-бот.
// Токен бота НЕ в коді: Налаштування проєкту → Властивості скрипту → BOT_TOKEN.
var SHEET_ID = '1Ao7XECDUYsqZKzbzvkwmcvZkr8LAwoo9Fciv22IKq4c';
var SITE = 'https://deniszshina-dot.github.io/ze-shina-feed/book/';
var PER_HOUR = 2;                       // машин на годину
var HOURS = {1:[9,16],2:[9,16],3:[9,16],4:[9,16],5:[9,16],6:[9,14]}; // день тижня → [перша, остання] година запису; неділя — вихідний
var LEAD_MIN = 120;                     // запис із сайту на сьогодні — не раніше ніж за 2 години
var TZ = 'Europe/Kyiv';
var DAYS_AHEAD = 60;
var TYRES = { storage: 'на зберіганні у нас', 'new': 'нові, куплені у нас (чекають монтажу)', bring: 'привезе з собою' };
var C = { id: 0, created: 1, date: 2, time: 3, plate: 4, car: 5, phone: 6, svc: 7, tyres: 8, src: 9, status: 10, stor: 11, msg: 12, client: 13, remind: 14 };
var HEAD = ['id', 'створено', 'дата', 'час', 'номер', 'авто', 'телефон', 'послуги', 'шини', 'джерело', 'статус', 'на зберіганні', 'tg_msg', 'tg_клієнт', 'нагадування'];

// ---------- HTTP ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.slots) return out_(slots_(p.slots, true));
  if (p.cfg) return out_({ bot: botName_() });
  var plate = normPlate_(p.plate || '');
  if (!/^[A-Z]{2}\d{4}[A-Z]{2}$/.test(plate)) return out_({ error: 'bad_plate' });
  if (!throttle_('get', 300)) return out_({ error: 'busy' });
  var st = storageFor_(plate);
  if (!st) return out_({ found: false });
  return out_({ found: true, unsure: st.unsure, tyres: st.unsure ? '' : st.tyres, qty: st.unsure ? '' : st.qty,
                discs: !st.unsure && st.discs, since: st.unsure ? '' : st.since });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad_json' }); }
  if (!throttle_('post', 30)) return out_({ ok: false, error: 'busy' });
  if (d.action === 'book') {
    var r = book_({ date: d.date, time: d.time, plate: d.plate, car: d.car, phone: d.phone,
                    services: d.services, tyres: d.tyres, source: 'сайт' });
    if (r.ok) r.bot = botName_();
    return out_(r);
  }
  if (d.action === 'cancel') return out_(clientCancel_(String(d.id || ''), String(d.phone || ''), null));
  return out_({ ok: false, error: 'bad_action' });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function throttle_(k, max) {
  var c = CacheService.getScriptCache(), n = Number(c.get(k) || 0);
  if (n > max) return false;
  c.put(k, String(n + 1), 60);
  return true;
}

// ---------- Зберігання ----------
function normPlate_(v) {
  var map = {'А':'A','В':'B','Е':'E','І':'I','К':'K','М':'M','Н':'H','О':'O','Р':'P','С':'C','Т':'T','Х':'X','Й':'Y'};
  return String(v).toUpperCase().split('').map(function (c) { return map[c] || c; }).join('').replace(/[^A-Z0-9]/g, '');
}
function digits9_(v) { return String(v || '').replace(/\D/g, '').replace(/^380/, '').replace(/^0/, '').slice(-9); }

function storageFor_(plate) {
  var rows = SpreadsheetApp.openById(SHEET_ID).getSheets()[0].getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    if (normPlate_(rows[i][0]) === plate) {
      var note = String(rows[i][6] || '').toLowerCase();
      return { unsure: /выдал|видал|забрал|забра|утилиз|продаж/.test(note), tyres: rows[i][1], qty: rows[i][2],
               discs: rows[i][3] === 'так', since: rows[i][4], note: rows[i][6] };
    }
  }
  return null;
}

// ---------- Таблиці ----------
function sheet_(name, head) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange('A:D').setNumberFormat('@');
  }
  return sh;
}
var BK_ = null;
function bookSheet_() {
  if (BK_) return BK_;
  var sh = sheet_('Записи', HEAD);
  if (!sh.getRange(1, 14).getDisplayValue()) sh.getRange(1, 14, 1, 2).setValues([[HEAD[13], HEAD[14]]]);
  return (BK_ = sh);
}
function closeSheet_() { return sheet_('Закрито', ['дата', 'з', 'до', 'причина', 'хто']); }
function waitSheet_() { return sheet_('Очікування', ['дата', 'chat_id', 'створено']); }
function rows_() { return bookSheet_().getDataRange().getDisplayValues(); }
function cancelled_(s) { return /^скасовано/.test(String(s)); }
function todayYmd_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function addDays_(ymd, n) {
  var p = ymd.split('-'), d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n, 12));
  return d.toISOString().slice(0, 10);
}
function findRow_(v, id) {
  for (var i = 1; i < v.length; i++) if (v[i][C.id] === id || v[i][C.id] === String(Number(id))) return i;
  return -1;
}

// ---------- Розклад ----------
function dayInfo_(ymd) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return null;
  var p = ymd.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12);
  var today = todayYmd_();
  var max = Utilities.formatDate(new Date(Date.now() + DAYS_AHEAD * 864e5), TZ, 'yyyy-MM-dd');
  if (ymd < today || ymd > max) return null;
  var h = HOURS[d.getDay()];
  if (!h) return { closed: true };
  return { from: h[0], to: h[1], dow: d.getDay(), today: ymd === today };
}

// Закриті години дня: {all: bool, h: {9:'причина', ...}}
function closedFor_(ymd) {
  var v = closeSheet_().getDataRange().getDisplayValues(), res = { all: false, h: {}, why: '' };
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] !== ymd) continue;
    if (!v[i][1]) { res.all = true; res.why = v[i][3]; continue; }
    for (var h = Number(v[i][1]); h <= Number(v[i][2]); h++) res.h[h] = v[i][3] || 'закрито';
  }
  return res;
}

function busyMap_(ymd) {
  var v = rows_(), m = {};
  for (var i = 1; i < v.length; i++) {
    if (v[i][C.date] === ymd && !cancelled_(v[i][C.status])) m[v[i][C.time]] = (m[v[i][C.time]] || 0) + 1;
  }
  return m;
}

function nowMin_() { return Number(Utilities.formatDate(new Date(), TZ, 'H')) * 60 + Number(Utilities.formatDate(new Date(), TZ, 'm')); }

function slots_(ymd, forSite) {
  var di = dayInfo_(ymd);
  if (!di) return { error: 'bad_date' };
  if (di.closed) return { closed: true, slots: [] };
  var cl = closedFor_(ymd);
  if (cl.all) return { closed: true, why: cl.why, slots: [] };
  var busy = busyMap_(ymd), now = nowMin_(), res = [];
  for (var h = di.from; h <= di.to; h++) {
    var t = (h < 10 ? '0' : '') + h + ':00';
    var left = Math.max(0, PER_HOUR - (busy[t] || 0));
    if (cl.h[h]) left = 0;
    if (di.today && h * 60 < now + (forSite ? LEAD_MIN : 0)) left = 0;
    res.push({ time: t, left: left });
  }
  return { slots: res };
}

function book_(b) {
  var site = b.source === 'сайт';
  var plate = normPlate_(b.plate || '');
  var phone = digits9_(b.phone);
  if ((plate || site) && !/^[ABCEHIKMOPTX]{2}\d{4}[A-Z]{2}$/.test(plate)) return { ok: false, error: 'bad_plate' };
  if (site && !/^(39|50|63|66|67|68|73|75|77|91|92|93|94|95|96|97|98|99)\d{7}$/.test(phone)) return { ok: false, error: 'bad_phone' };
  var di = dayInfo_(String(b.date || ''));
  if (!di || di.closed) return { ok: false, error: 'bad_date' };
  var t = String(b.time || '');
  var h = Number(t.slice(0, 2));
  if (!/^\d{2}:00$/.test(t) || h < di.from || h > di.to) return { ok: false, error: 'bad_time' };
  var cl = closedFor_(b.date);
  if (cl.all || cl.h[h]) return { ok: false, error: 'closed' };
  if (site && di.today && h * 60 < nowMin_() + LEAD_MIN) return { ok: false, error: 'too_soon' };

  var lock = LockService.getScriptLock(), done = false;
  lock.waitLock(20000);
  try {
    var busy = busyMap_(b.date);
    if ((busy[t] || 0) >= PER_HOUR) return { ok: false, error: 'full' };
    var sh = bookSheet_();
    var plateF = plate ? plate.slice(0, 2) + ' ' + plate.slice(2, 6) + ' ' + plate.slice(6) : '';
    var phoneF = phone.length === 9 ? '0' + phone.slice(0, 2) + ' ' + phone.slice(2, 5) + ' ' + phone.slice(5, 7) + ' ' + phone.slice(7) : '';
    var ex = sh.getDataRange().getDisplayValues(), today = todayYmd_(), noShow = 0;
    for (var k = 1; k < ex.length; k++) {
      var r = ex[k], same = (plate && r[C.plate] === plateF) || (phone && digits9_(r[C.phone]) === phone);
      if (!same) continue;
      if (r[C.status] === 'не приїхав') noShow++;
      if (cancelled_(r[C.status])) continue;
      if (plate && r[C.date] === b.date && r[C.plate] === plateF) return { ok: false, error: 'dup' };
      if (site && r[C.date] >= today && r[C.status] !== 'приїхав' && r[C.status] !== 'не приїхав')
        return { ok: false, error: 'has_booking', date: r[C.date], time: r[C.time] };
    }
    // унікальний id з літерою (таблиця не перетворить на число); перевіряємо, що такого ще немає
    var id, taken = {};
    for (var q = 1; q < ex.length; q++) taken[ex[q][C.id]] = 1;
    var seed = Date.now();
    do { id = 'z' + (seed++).toString(36); } while (taken[id]);
    var st = plate ? storageFor_(plate) : null;
    var stTxt = st ? (st.unsure ? 'є запис, примітка: ' + st.note : st.tyres + (st.discs ? ' + диски' : '') + (st.note ? ' [' + st.note + ']' : '')) : '';
    var services = String(b.services || '').slice(0, 200);
    var tyres = TYRES[b.tyres] || String(b.tyres || '').slice(0, 60);
    var car = String(b.car || '').slice(0, 120);
    sh.appendRow([id, Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'), b.date, t, plateF, car, phoneF, services, tyres, b.source, 'нова', stTxt, '', '', '']);
    var row = sh.getLastRow();
    var left = PER_HOUR - (busy[t] || 0) - 1;
    var msg = tg_('sendMessage', {
      chat_id: chatId_(), parse_mode: 'HTML',
      text: '🆕 <b>Новий запис</b> (' + esc_(b.source) + ')\n' + niceDate_(b.date) + ' · <b>' + t + '</b> (вільних місць на цю годину: ' + left + ')\n\n' +
            '🚗 <b>' + (plateF || 'без номера') + '</b>' + (car ? ' · ' + esc_(car) : '') + '\n' +
            (phoneF ? '📞 ' + phoneF + '\n' : '') +
            (services ? '🔧 ' + esc_(services) + '\n' : '') +
            '📦 Шини: ' + esc_(tyres) + (stTxt ? '\n🗄 На зберіганні: ' + esc_(stTxt) : '') +
            (noShow ? '\n⚠️ Раніше не приїжджав: ' + noShow + ' раз(и)' : ''),
      reply_markup: kb_(id)
    });
    if (msg && msg.ok) sh.getRange(row, C.msg + 1).setValue(msg.result.message_id);
    done = true;
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
    if (done) refreshPin_();
  }
}

function kb_(id) {
  return JSON.stringify({ inline_keyboard: [[
    { text: '🚗 Приїхав', callback_data: 'ar:' + id }, { text: '🚫 Не приїхав', callback_data: 'ns:' + id },
    { text: '❌ Скасувати', callback_data: 'no:' + id }]] });
}

// Скасування: статус, повідомлення в чат, клієнту, черга очікування, закріп.
function cancelRow_(i, v, status, who) {
  var sh = bookSheet_(), r = v[i];
  sh.getRange(i + 1, C.status + 1).setValue(status);
  var line = niceDate_(r[C.date]) + ' ' + r[C.time] + ' — ' + ((r[C.plate] ? r[C.plate] + ' ' : '') + r[C.car]).slice(0, 120);
  if (r[C.msg]) tg_('editMessageReplyMarkup', { chat_id: chatId_(), message_id: r[C.msg], reply_markup: JSON.stringify({ inline_keyboard: [] }) });
  if (status === 'скасовано клієнтом') {
    tg_('sendMessage', { chat_id: chatId_(), text: '❌ Клієнт сам скасував запис: ' + line + (r[C.phone] ? '\n📞 ' + r[C.phone] : '') + '\nМісце звільнилось.' });
  } else if (r[C.client]) {
    tg_('sendMessage', { chat_id: r[C.client], text: 'Ваш запис на шиномонтаж ZE-Shina ' + niceDate_(r[C.date]) + ' о ' + r[C.time] + ' скасовано. Якщо це помилка — подзвоніть: 099 223 25 52.\nЗаписатися знову: ' + SITE });
  }
  notifyWait_(r[C.date]);
  return line;
}

function clientCancel_(id, phone, chat) {
  var v = rows_(), i = findRow_(v, id);
  if (i < 0) return { ok: false, error: 'not_found' };
  var r = v[i];
  var okWho = chat ? String(r[C.client]) === String(chat) : (digits9_(phone).length === 9 && digits9_(r[C.phone]) === digits9_(phone));
  if (!okWho) return { ok: false, error: 'not_found' };
  if (cancelled_(r[C.status])) return { ok: true, already: true };
  if (r[C.date] < todayYmd_()) return { ok: false, error: 'past' };
  cancelRow_(i, v, 'скасовано клієнтом');
  refreshPin_();
  return { ok: true };
}

function notifyWait_(ymd) {
  if (ymd < todayYmd_()) return;
  var sl = slots_(ymd, true);
  var free = (sl.slots || []).filter(function (x) { return x.left > 0; }).map(function (x) { return x.time; });
  if (!free.length) return;
  var sh = waitSheet_(), v = sh.getDataRange().getDisplayValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (v[i][0] !== ymd) continue;
    tg_('sendMessage', { chat_id: v[i][1], text: '🔔 Звільнилось місце на шиномонтаж ZE-Shina: ' + niceDate_(ymd) + ', вільно: ' + free.join(', ') + '.\nЗаписатися: ' + SITE + '?d=' + ymd });
    sh.deleteRow(i + 1);
  }
}

// ---------- Telegram ----------
function token_() { return PropertiesService.getScriptProperties().getProperty('BOT_TOKEN'); }
function chatId_() { return PropertiesService.getScriptProperties().getProperty('CHAT_ID'); }
function esc_(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function tg_(method, payload) {
  var tk = token_();
  if (!tk) return null;
  var r = UrlFetchApp.fetch('https://api.telegram.org/bot' + tk + '/' + method,
    { method: 'post', payload: payload, muteHttpExceptions: true });
  try { return JSON.parse(r.getContentText()); } catch (x) { return null; }
}

function botName_() {
  var props = PropertiesService.getScriptProperties(), n = props.getProperty('BOT_NAME');
  if (n) return n;
  var r = tg_('getMe', {});
  if (r && r.ok) { n = r.result.username; props.setProperty('BOT_NAME', n); }
  return n || '';
}

function niceDate_(ymd) {
  var p = ymd.split('-'), d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12);
  var dn = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'][d.getDay()];
  return dn + ', ' + p[2] + '.' + p[1];
}

// Запускається тригером щохвилини: читає натискання кнопок і команди в чаті.
function poll() {
  var props = PropertiesService.getScriptProperties();
  var off = Number(props.getProperty('OFFSET') || 0);
  var r = tg_('getUpdates', { offset: off, timeout: 0, allowed_updates: JSON.stringify(['message', 'callback_query']) });
  if (r && r.ok) {
    r.result.forEach(function (u) {
      off = u.update_id + 1;
      try { handle_(u); } catch (x) { console.error(x); }
    });
    props.setProperty('OFFSET', String(off));
  }
  if (props.getProperty('PIN_DAY') !== todayYmd_() && Number(Utilities.formatDate(new Date(), TZ, 'H')) >= 6) refreshPin_();
}

function handle_(u) {
  var props = PropertiesService.getScriptProperties();
  if (u.callback_query) return handleCb_(u.callback_query);
  var m = u.message;
  if (!m || !m.text) return;
  var text = m.text.trim();
  if (m.chat.type === 'private') return handlePrivate_(m, text);
  // Прив'язка чату: написати в потрібній групі /bind
  if (/^\/bind(@\w+)?$/.test(text)) {
    props.setProperty('CHAT_ID', String(m.chat.id));
    props.deleteProperty('PIN_ID');
    tg_('sendMessage', { chat_id: m.chat.id, text: 'Готово: записи будуть приходити в цей чат.' });
    refreshPin_();
    return;
  }
  if (String(m.chat.id) !== String(chatId_())) return;
  if (/^\/(plan|план)/i.test(text)) { sendPlan_(m.chat.id, text.split(/\s+/)[1]); return; }
  if (/^\/(zapis|запис)/i.test(text)) { cmdBook_(m, text); return; }
  if (/^\/(cancel|otmena|скасувати|отмена)/i.test(text)) { cmdCancel_(m, text); return; }
  if (/^\/close/i.test(text)) { cmdClose_(m, text, true); return; }
  if (/^\/open/i.test(text)) { cmdClose_(m, text, false); return; }
  if (/^\/pin/i.test(text)) { props.deleteProperty('PIN_ID'); refreshPin_(); return; }
  if (/^\/(help|start)/i.test(text)) { tg_('sendMessage', { chat_id: m.chat.id, text: HELP_ }); return; }
}

var HELP_ = 'Команди:\n' +
  '/zapis 28.11 10 AX1234KA Шкода 0671234567 збер — записати (номер і телефон — якщо є; шини: збер / нові / без слова = свої)\n' +
  '/cancel 28.11 10 AX1234KA — скасувати (або слово із запису)\n' +
  '/plan 28.11 — список на день\n' +
  '/close 28.11 — закрити день;  /close 28.11 14-16 світло — закрити години\n' +
  '/open 28.11 — відкрити знову\n' +
  '/pin — створити закріплений розклад заново';

function handleCb_(q) {
  var parts = String(q.data).split(':'), act = parts[0], id = parts[1];
  var chat = String(q.message.chat.id);
  var v = rows_(), i = findRow_(v, id);
  // Кнопки клієнта в особистому чаті
  if (act === 'cy' || act === 'cc') {
    if (i < 0 || String(v[i][C.client]) !== chat) { tg_('answerCallbackQuery', { callback_query_id: q.id, text: 'Запис не знайдено' }); return; }
    if (act === 'cy') {
      bookSheet_().getRange(i + 1, C.remind + 1).setValue('підтвердив');
      tg_('editMessageText', { chat_id: chat, message_id: q.message.message_id, text: q.message.text + '\n\n✅ Дякуємо, чекаємо на вас!' });
      refreshPin_();
    } else {
      var r = clientCancel_(id, '', chat);
      tg_('editMessageText', { chat_id: chat, message_id: q.message.message_id, text: q.message.text + '\n\n' + (r.ok ? '❌ Запис скасовано. Дякуємо, що попередили!' : 'Не вдалося скасувати — подзвоніть: 099 223 25 52') });
    }
    tg_('answerCallbackQuery', { callback_query_id: q.id });
    return;
  }
  if (chat !== String(chatId_())) return;
  var who = q.from.first_name || '';
  var st = { ar: 'приїхав', ns: 'не приїхав', no: 'скасовано', ok: 'прийнято' }[act];
  if (!st) return;
  var tail = { ar: '🚗 Приїхав', ns: '🚫 Не приїхав', no: '❌ Скасовано', ok: '✅ Прийнято' }[act] + ' — ' + esc_(who);
  if (i >= 0) {
    if (st === 'скасовано') cancelRow_(i, v, 'скасовано', who);
    else bookSheet_().getRange(i + 1, C.status + 1).setValue(st);
  }
  tg_('editMessageText', { chat_id: chat, message_id: q.message.message_id, parse_mode: 'HTML', text: esc_(q.message.text) + '\n\n' + tail,
                           reply_markup: st === 'скасовано' ? JSON.stringify({ inline_keyboard: [] }) : kb_(id) });
  tg_('answerCallbackQuery', { callback_query_id: q.id, text: st });
  refreshPin_();
}

function handlePrivate_(m, text) {
  var arg = (text.match(/^\/start\s+(\S+)/) || [])[1] || '';
  if (/^w\d{8}$/.test(arg)) {
    var ymd = arg.slice(1, 5) + '-' + arg.slice(5, 7) + '-' + arg.slice(7);
    waitSheet_().appendRow([ymd, String(m.chat.id), Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm')]);
    tg_('sendMessage', { chat_id: m.chat.id, text: '🔔 Добре! Якщо на ' + niceDate_(ymd) + ' звільниться місце, ми одразу напишемо сюди.' });
    return;
  }
  if (/^z[0-9a-z]{6,12}$/.test(arg)) {
    var v = rows_(), i = findRow_(v, arg);
    if (i >= 0 && !cancelled_(v[i][C.status])) {
      bookSheet_().getRange(i + 1, C.client + 1).setValue(String(m.chat.id));
      var r = v[i];
      tg_('sendMessage', { chat_id: m.chat.id,
        text: '✅ Ваш запис на шиномонтаж ZE-Shina:\n' + niceDate_(r[C.date]) + ' о ' + r[C.time] + '\n' + (r[C.plate] ? r[C.plate] + '\n' : '') + (r[C.svc] ? r[C.svc] + '\n' : '') +
              'вул. Шевченка, 24\n\nНапередодні ввечері нагадаємо сюди. Не забудьте ключ від секреток, якщо вони стоять.',
        reply_markup: JSON.stringify({ inline_keyboard: [[{ text: '❌ Скасувати запис', callback_data: 'cc:' + r[C.id] }]] }) });
      return;
    }
  }
  tg_('sendMessage', { chat_id: m.chat.id, text: 'Це бот запису на шиномонтаж ZE-Shina (Харків, вул. Шевченка, 24).\nЗаписатися: ' + SITE });
}

var PLATE_RE = /(^|[^A-Za-zА-Яа-яІіЇїЙй0-9])([A-Za-zА-Яа-яІіЇїЙй]{2})\s?(\d{4})\s?([A-Za-zА-Яа-яІіЇїЙй]{2})(?=$|[^A-Za-zА-Яа-яІіЇїЙй0-9])/;
var PHONE_RE = /(?:\+?38\s?)?0(?:[\s-]?\d){9}(?!\d)/;

// "28.11 10 ..." → {ymd, time, rest}
function parseDT_(args) {
  var mm = args.match(/^(\d{1,2})\.(\d{1,2})\s+(\d{1,2})(?:[:.-]\d{2})?(?:\s+([\s\S]*))?$/);
  if (!mm) return null;
  return { ymd: ymd_(mm[1], mm[2]), time: ('0' + mm[3]).slice(-2) + ':00', rest: (mm[4] || '').trim() };
}
function ymd_(dd, mm) {
  var y = Number(Utilities.formatDate(new Date(), TZ, 'yyyy'));
  var ymd = y + '-' + ('0' + mm).slice(-2) + '-' + ('0' + dd).slice(-2);
  if (ymd < todayYmd_()) ymd = (y + 1) + ymd.slice(4);
  return ymd;
}

function cmdBook_(m, text) {
  var dt = parseDT_(text.replace(/^\/\S+\s*/, ''));
  var help = 'Формат: /zapis 28.11 10 AX1234KA Шкода Октавія 0671234567 збер\n' +
             'Номер авто і телефон — якщо є. Шини: збер (зі зберігання) / нові / без слова = привезе свої.';
  if (!dt || !dt.rest) { tg_('sendMessage', { chat_id: m.chat.id, text: help }); return; }
  var tail = dt.rest, plate = '', phone = '';
  var pm = tail.match(PLATE_RE);
  if (pm) { plate = pm[2] + pm[3] + pm[4]; tail = tail.replace(pm[0], pm[1] + ' '); }
  var ph = tail.match(PHONE_RE);
  if (ph) { phone = ph[0]; tail = tail.replace(ph[0], ' '); }
  var ty = /збер|хран/i.test(tail) ? 'storage' : /нов|оплач/i.test(tail) ? 'new' : 'bring';
  var car = tail.replace(/(^|\s)(своя|свої|свои|збер\S*|хран\S*)(?=\s|$)/ig, ' ').replace(/\s+/g, ' ').trim();
  if (!plate && !car) { tg_('sendMessage', { chat_id: m.chat.id, text: help }); return; }
  var res = book_({ date: dt.ymd, time: dt.time, plate: plate, car: car, phone: phone,
                    services: '', tyres: ty, source: 'чат (' + (m.from.first_name || '') + ')' });
  if (!res.ok) tg_('sendMessage', { chat_id: m.chat.id, text: 'Не записав: ' + ({ full: 'на цю годину вже 2 машини', bad_plate: 'номер авто неповний — напишіть повністю або приберіть', bad_date: 'дата поза графіком (неділя або минула дата)', bad_time: 'час поза графіком (Пн–Пт 9–16, Сб 9–14)', dup: 'це авто вже записане на цей день', closed: 'цю годину закрито (/open, щоб відкрити)' }[res.error] || res.error) });
}

function cmdCancel_(m, text) {
  var dt = parseDT_(text.replace(/^\/\S+\s*/, ''));
  var help = 'Формат: /cancel 28.11 10 AX1234KA\n(дата, година і номер авто або слово із запису, напр. Шкода)';
  if (!dt || !dt.rest) { tg_('sendMessage', { chat_id: m.chat.id, text: help }); return; }
  var key = normPlate_(dt.rest), word = dt.rest.toLowerCase();
  var v = rows_(), hit = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i][C.date] !== dt.ymd || v[i][C.time] !== dt.time || cancelled_(v[i][C.status])) continue;
    var pl = v[i][C.plate].replace(/\s/g, '');
    if ((key.length >= 4 && pl && pl.indexOf(key) >= 0) || v[i][C.car].toLowerCase().indexOf(word) >= 0) hit.push(i);
  }
  if (!hit.length) { tg_('sendMessage', { chat_id: m.chat.id, text: 'Не знайшов такого запису на ' + niceDate_(dt.ymd) + ' ' + dt.time + '. Перевірте: /plan ' + dt.ymd.slice(8) + '.' + dt.ymd.slice(5, 7) }); return; }
  if (hit.length > 1) { tg_('sendMessage', { chat_id: m.chat.id, text: 'Підходить кілька записів, уточніть:\n' + hit.map(function (i) { return '• ' + (v[i][C.plate] || '') + ' ' + v[i][C.car]; }).join('\n') }); return; }
  var line = cancelRow_(hit[0], v, 'скасовано', m.from.first_name || '');
  tg_('sendMessage', { chat_id: m.chat.id, text: '❌ Скасовано: ' + line + '\nМісце звільнилось.' });
  refreshPin_();
}

// /close 28.11 [14-16] [причина]   /open 28.11 [14-16]
function cmdClose_(m, text, close) {
  var a = text.replace(/^\/\S+\s*/, '').match(/^(\d{1,2})\.(\d{1,2})(?:\s+(\d{1,2})(?:\s*-\s*(\d{1,2}))?)?\s*(.*)$/);
  if (!a) { tg_('sendMessage', { chat_id: m.chat.id, text: 'Формат: /close 28.11  або  /close 28.11 14-16 причина\n/open 28.11' }); return; }
  var ymd = ymd_(a[1], a[2]), from = a[3] ? Number(a[3]) : null, to = a[4] ? Number(a[4]) : from, why = (a[5] || '').trim();
  var sh = closeSheet_();
  if (!close) {
    var v = sh.getDataRange().getDisplayValues(), n = 0;
    for (var i = v.length - 1; i >= 1; i--) {
      if (v[i][0] === ymd && (from === null || (Number(v[i][1]) === from && Number(v[i][2]) === to))) { sh.deleteRow(i + 1); n++; }
    }
    tg_('sendMessage', { chat_id: m.chat.id, text: n ? '✅ Відкрито: ' + niceDate_(ymd) + (from !== null ? ' ' + from + '–' + to : '') : 'Закриттів на ' + niceDate_(ymd) + ' не знайшов.' });
    notifyWait_(ymd);
    refreshPin_();
    return;
  }
  sh.appendRow([ymd, from === null ? '' : String(from), from === null ? '' : String(to), why, m.from.first_name || '']);
  var bk = rows_().filter(function (r, k) {
    if (!k || r[C.date] !== ymd || cancelled_(r[C.status])) return false;
    var hh = Number(r[C.time].slice(0, 2));
    return from === null || (hh >= from && hh <= to);
  });
  tg_('sendMessage', { chat_id: m.chat.id, text: '⛔ Закрито для запису: ' + niceDate_(ymd) + (from !== null ? ' ' + from + ':00–' + to + ':00' : ' (весь день)') + (why ? ' — ' + why : '') +
    (bk.length ? '\n\nУже записані на цей час (зателефонуйте їм):\n' + bk.map(function (r) { return r[C.time] + ' ' + (r[C.plate] || '') + ' ' + r[C.car].slice(0, 50) + (r[C.phone] ? ' 📞 ' + r[C.phone] : ''); }).join('\n') : '') });
  refreshPin_();
}

// ---------- План, склад, нагадування, закріп ----------
function dayRows_(ymd) {
  return rows_().slice(1).filter(function (r) { return r[C.date] === ymd && !cancelled_(r[C.status]); })
    .sort(function (a, b) { return a[C.time] < b[C.time] ? -1 : a[C.time] > b[C.time] ? 1 : 0; });
}
function tyShort_(s) { return /зберіг/.test(s) ? 'збер' : /нов/.test(s) ? 'нові' : /привез/.test(s) ? 'свої' : ''; }

function sendPlan_(chat, arg) {
  var ymd = arg && /^\d{1,2}\.\d{1,2}$/.test(arg) ? ymd_(arg.split('.')[0], arg.split('.')[1]) : addDays_(todayYmd_(), 1);
  var v = dayRows_(ymd);
  var s = v.length ? v.map(function (r) { return r[C.time] + '  ' + r[C.plate] + (r[C.car] ? ' · ' + r[C.car] : '') + '  — шини: ' + r[C.tyres] + (r[C.stor] ? '\n        🗄 ' + r[C.stor] : ''); }).join('\n') : 'записів немає';
  var sl = slots_(ymd, false);
  var free = sl.slots ? sl.slots.reduce(function (a, x) { return a + x.left; }, 0) : 0;
  tg_('sendMessage', { chat_id: chat, text: '📋 ' + niceDate_(ymd) + ' — ' + v.length + ' маш.\n' + s + '\n\nВільних місць: ' + free + (sl.closed ? ' (день закрито)' : '') });
}

function dayBlock_(ymd, title) {
  var di = dayInfo_(ymd), cl = closedFor_(ymd), v = dayRows_(ymd);
  var out = '━━ ' + title + ', ' + niceDate_(ymd) + ' ━━\n';
  if (!di || di.closed) return out + 'вихідний\n';
  if (cl.all) out += '⛔ день закрито' + (cl.why ? ' — ' + cl.why : '') + '\n';
  var busy = {};
  v.forEach(function (r) {
    busy[r[C.time]] = (busy[r[C.time]] || 0) + 1;
    var mark = r[C.status] === 'приїхав' ? '🚗 ' : r[C.status] === 'не приїхав' ? '🚫 ' : r[C.remind] === 'підтвердив' ? '✅ ' : '';
    out += r[C.time].slice(0, 5) + ' ' + mark + (r[C.plate] ? r[C.plate].replace(/\s/g, '') + ' ' : '') + r[C.car].slice(0, 38) + (tyShort_(r[C.tyres]) ? ' · ' + tyShort_(r[C.tyres]) : '') + '\n';
  });
  var free = [];
  for (var h = di.from; h <= di.to; h++) {
    var t = (h < 10 ? '0' : '') + h + ':00';
    if (cl.all) break;
    if (cl.h[h]) { free.push(t.slice(0, 2) + '⛔'); continue; }
    var left = PER_HOUR - (busy[t] || 0);
    if (left > 0) free.push(t.slice(0, 2) + (left === PER_HOUR ? '×2' : '×1'));
  }
  if (!cl.all) out += free.length ? 'Вільно: ' + free.join(' ') + '\n' : 'Вільних місць немає\n';
  return out;
}

function refreshPin_() {
  var chat = chatId_();
  if (!chat || !token_()) return;
  var props = PropertiesService.getScriptProperties(), today = todayYmd_();
  var text = '📌 РОЗКЛАД — оновлюється автоматично\n\n' + dayBlock_(today, 'Сьогодні') + '\n' + dayBlock_(addDays_(today, 1), 'Завтра') +
             '\n🚗 приїхав · 🚫 не приїхав · ✅ клієнт підтвердив\nКоманди: /help';
  if (text.length > 4000) text = text.slice(0, 3990) + '…';
  var pid = props.getProperty('PIN_ID'), r = null;
  if (pid) {
    r = tg_('editMessageText', { chat_id: chat, message_id: pid, text: text });
    if (r && !r.ok && !/not modified/.test(r.description || '')) pid = null;
  }
  if (!pid) {
    r = tg_('sendMessage', { chat_id: chat, text: text, disable_notification: true });
    if (r && r.ok) {
      pid = String(r.result.message_id);
      props.setProperty('PIN_ID', pid);
      var p = tg_('pinChatMessage', { chat_id: chat, message_id: pid, disable_notification: true });
      if (p && !p.ok) tg_('sendMessage', { chat_id: chat, text: 'Не можу закріпити розклад: дайте боту право «Закріплювати повідомлення» (зробіть адміном) і напишіть /pin.' });
    }
  }
  props.setProperty('PIN_DAY', today);
}

// Запускається тригером щодня о 19:00: план на завтра, список для складу, нагадування клієнтам.
function dailyPlan() {
  if (!chatId_()) return;
  var tmr = addDays_(todayYmd_(), 1);
  sendPlan_(chatId_());
  var st = dayRows_(tmr).filter(function (r) { return /зберіг/.test(r[C.tyres]); });
  if (st.length) tg_('sendMessage', { chat_id: chatId_(), text: '📦 Винести зі складу на завтра (' + niceDate_(tmr) + '):\n' +
    st.map(function (r) { return r[C.time] + ' ' + (r[C.plate] || '') + ' — ' + (r[C.stor] || r[C.car]).slice(0, 110); }).join('\n') });
  var sh = bookSheet_(), v = rows_();
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    if (r[C.date] !== tmr || cancelled_(r[C.status]) || !r[C.client] || r[C.remind]) continue;
    var res = tg_('sendMessage', { chat_id: r[C.client],
      text: '🔔 Нагадуємо: завтра, ' + niceDate_(r[C.date]) + ' о ' + r[C.time] + ' — шиномонтаж ZE-Shina, вул. Шевченка, 24.\n' + (r[C.plate] ? r[C.plate] + '\n' : '') +
            'Не забудьте ключ від секреток, якщо вони стоять.\nПід час повітряної тривоги роботу призупиняємо.',
      reply_markup: JSON.stringify({ inline_keyboard: [[{ text: '✅ Буду', callback_data: 'cy:' + r[C.id] }, { text: '❌ Скасувати', callback_data: 'cc:' + r[C.id] }]] }) });
    sh.getRange(i + 1, C.remind + 1).setValue(res && res.ok ? 'нагадано' : 'помилка');
  }
  refreshPin_();
}

// Один раз вручну: створює тригери poll (щохвилини) і dailyPlan (19:00).
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('poll').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('dailyPlan').timeBased().atHour(19).everyDays(1).inTimezone(TZ).create();
  bookSheet_();
  tg_('deleteWebhook', {});
}

function test_() {
  Logger.log(doGet({ parameter: { plate: 'AX6486YA' } }).getContent());
  Logger.log(JSON.stringify(slots_(addDays_(todayYmd_(), 1), true)));
}
