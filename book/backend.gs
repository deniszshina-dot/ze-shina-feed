// ZE-Shina: запис на шиномонтаж + пошук зберігання + Telegram-бот.
// Токен бота НЕ в коді: Налаштування проєкту → Властивості скрипту → BOT_TOKEN.
var SHEET_ID = '1Ao7XECDUYsqZKzbzvkwmcvZkr8LAwoo9Fciv22IKq4c';
var PER_HOUR = 2;                       // машин на годину
var HOURS = {1:[9,16],2:[9,16],3:[9,16],4:[9,16],5:[9,16],6:[9,14]}; // день тижня → [перша, остання] година запису; неділя — вихідний
var TZ = 'Europe/Kyiv';
var DAYS_AHEAD = 60;
var TYRES = { storage: 'на зберіганні у нас', 'new': 'нові, куплені у нас (чекають монтажу)', bring: 'привезе з собою' };

// ---------- HTTP ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.slots) return out_(slots_(p.slots));
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
  if (d.action !== 'book') return out_({ ok: false, error: 'bad_action' });
  if (!throttle_('post', 30)) return out_({ ok: false, error: 'busy' });
  var r = book_({ date: d.date, time: d.time, plate: d.plate, car: d.car, phone: d.phone,
                  services: d.services, tyres: d.tyres, source: 'сайт' });
  return out_(r);
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

// ---------- Записи ----------
function bookSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Записи');
  if (!sh) {
    sh = ss.insertSheet('Записи');
    sh.appendRow(['id', 'створено', 'дата', 'час', 'номер', 'авто', 'телефон', 'послуги', 'шини', 'джерело', 'статус', 'на зберіганні', 'tg_msg']);
    sh.setFrozenRows(1);
    sh.getRange('C:D').setNumberFormat('@');
  }
  return sh;
}

function dayInfo_(ymd) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return null;
  var p = ymd.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12);
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var max = Utilities.formatDate(new Date(Date.now() + DAYS_AHEAD * 864e5), TZ, 'yyyy-MM-dd');
  if (ymd < today || ymd > max) return null;
  var h = HOURS[d.getDay()];
  if (!h) return { closed: true };
  return { from: h[0], to: h[1], dow: d.getDay(), today: ymd === today };
}

function busyMap_(ymd) {
  var v = bookSheet_().getDataRange().getDisplayValues(), m = {};
  for (var i = 1; i < v.length; i++) {
    if (v[i][2] === ymd && v[i][10] !== 'скасовано') m[v[i][3]] = (m[v[i][3]] || 0) + 1;
  }
  return m;
}

function slots_(ymd) {
  var di = dayInfo_(ymd);
  if (!di) return { error: 'bad_date' };
  if (di.closed) return { closed: true, slots: [] };
  var busy = busyMap_(ymd), nowH = Number(Utilities.formatDate(new Date(), TZ, 'H')), res = [];
  for (var h = di.from; h <= di.to; h++) {
    var t = (h < 10 ? '0' : '') + h + ':00';
    var left = Math.max(0, PER_HOUR - (busy[t] || 0));
    if (di.today && h <= nowH) left = 0;
    res.push({ time: t, left: left });
  }
  return { slots: res };
}

function book_(b) {
  var plate = normPlate_(b.plate || '');
  var phone = String(b.phone || '').replace(/\D/g, '').replace(/^380/, '').replace(/^0/, '');
  if ((plate || b.source === 'сайт') && !/^[ABCEHIKMOPTX]{2}\d{4}[A-Z]{2}$/.test(plate)) return { ok: false, error: 'bad_plate' };
  if (b.source === 'сайт' && !/^(39|50|63|66|67|68|73|75|77|91|92|93|94|95|96|97|98|99)\d{7}$/.test(phone)) return { ok: false, error: 'bad_phone' };
  var di = dayInfo_(String(b.date || ''));
  if (!di || di.closed) return { ok: false, error: 'bad_date' };
  var t = String(b.time || '');
  var h = Number(t.slice(0, 2));
  if (!/^\d{2}:00$/.test(t) || h < di.from || h > di.to) return { ok: false, error: 'bad_time' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var busy = busyMap_(b.date);
    if ((busy[t] || 0) >= PER_HOUR) return { ok: false, error: 'full' };
    var sh = bookSheet_();
    var plateF0 = plate ? plate.slice(0, 2) + ' ' + plate.slice(2, 6) + ' ' + plate.slice(6) : '';
    var ex = sh.getDataRange().getDisplayValues();
    for (var k = 1; plate && k < ex.length; k++) {
      if (ex[k][2] === b.date && ex[k][4] === plateF0 && ex[k][10] !== 'скасовано') return { ok: false, error: 'dup' };
    }
    var id = String(Date.now()).slice(-6);
    var st = plate ? storageFor_(plate) : null;
    var stTxt = st ? (st.unsure ? 'є запис, примітка: ' + st.note : st.tyres + (st.discs ? ' + диски' : '') + (st.note ? ' [' + st.note + ']' : '')) : '';
    var services = String(b.services || '').slice(0, 200);
    var tyres = TYRES[b.tyres] || String(b.tyres || '').slice(0, 60);
    var car = String(b.car || '').slice(0, 120);
    var plateF = plateF0;
    var phoneF = phone.length === 9 ? '0' + phone.slice(0, 2) + ' ' + phone.slice(2, 5) + ' ' + phone.slice(5, 7) + ' ' + phone.slice(7) : '';
    sh.appendRow([id, Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'), b.date, t, plateF, car, phoneF, services, tyres, b.source, 'нова', stTxt, '']);
    var row = sh.getLastRow();
    var left = PER_HOUR - (busy[t] || 0) - 1;
    var msg = tg_('sendMessage', {
      chat_id: chatId_(), parse_mode: 'HTML',
      text: '🆕 <b>Новий запис</b> (' + esc_(b.source) + ')\n' + niceDate_(b.date) + ' · <b>' + t + '</b> (вільних місць на цю годину: ' + left + ')\n\n' +
            '🚗 <b>' + (plateF || 'без номера') + '</b>' + (car ? ' · ' + esc_(car) : '') + '\n' +
            (phoneF ? '📞 ' + phoneF + '\n' : '') +
            (services ? '🔧 ' + esc_(services) + '\n' : '') +
            '📦 Шини: ' + esc_(tyres) + (stTxt ? '\n🗄 На зберіганні: ' + esc_(stTxt) : ''),
      reply_markup: JSON.stringify({ inline_keyboard: [[
        { text: '✅ Прийнято', callback_data: 'ok:' + id }, { text: '❌ Скасувати', callback_data: 'no:' + id }]] })
    });
    if (msg && msg.ok) sh.getRange(row, 13).setValue(msg.result.message_id);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
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
  if (!r || !r.ok) return;
  r.result.forEach(function (u) {
    off = u.update_id + 1;
    try { handle_(u); } catch (x) { console.error(x); }
  });
  props.setProperty('OFFSET', String(off));
}

function handle_(u) {
  var props = PropertiesService.getScriptProperties();
  if (u.callback_query) {
    var q = u.callback_query;
    if (String(q.message.chat.id) !== String(chatId_())) return;
    var parts = q.data.split(':'), st = parts[0] === 'ok' ? 'прийнято' : 'скасовано';
    var sh = bookSheet_(), v = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < v.length; i++) {
      if (v[i][0] === parts[1]) { sh.getRange(i + 1, 11).setValue(st); break; }
    }
    var who = q.from.first_name || '';
    tg_('editMessageText', { chat_id: q.message.chat.id, message_id: q.message.message_id, parse_mode: 'HTML',
      text: esc_(q.message.text) + '\n\n' + (st === 'прийнято' ? '✅ Прийнято' : '❌ Скасовано') + ' — ' + esc_(who) });
    tg_('answerCallbackQuery', { callback_query_id: q.id, text: st });
    return;
  }
  var m = u.message;
  if (!m || !m.text) return;
  var text = m.text.trim();
  // Прив'язка чату: написати в потрібній групі /bind
  if (/^\/bind(@\w+)?$/.test(text)) {
    props.setProperty('CHAT_ID', String(m.chat.id));
    tg_('sendMessage', { chat_id: m.chat.id, text: 'Готово: записи будуть приходити в цей чат.' });
    return;
  }
  if (String(m.chat.id) !== String(chatId_())) return;
  if (/^\/(plan|план)/i.test(text)) { sendPlan_(m.chat.id, text.split(/\s+/)[1]); return; }
  if (/^\/(zapis|запис)/i.test(text)) { cmdBook_(m, text); return; }
  if (/^\/(cancel|otmena|скасувати|отмена)/i.test(text)) { cmdCancel_(m, text); return; }
}

var PLATE_RE = /(^|[^A-Za-zА-Яа-яІіЇїЙй0-9])([A-Za-zА-Яа-яІіЇїЙй]{2})\s?(\d{4})\s?([A-Za-zА-Яа-яІіЇїЙй]{2})(?=$|[^A-Za-zА-Яа-яІіЇїЙй0-9])/;
var PHONE_RE = /(?:\+?38\s?)?0(?:[\s-]?\d){9}(?!\d)/;

// "28.11 10 ..." → {ymd, time, rest}
function parseDT_(args) {
  var mm = args.match(/^(\d{1,2})\.(\d{1,2})\s+(\d{1,2})(?:[:.-]\d{2})?(?:\s+([\s\S]*))?$/);
  if (!mm) return null;
  var y = Number(Utilities.formatDate(new Date(), TZ, 'yyyy'));
  var ymd = y + '-' + ('0' + mm[2]).slice(-2) + '-' + ('0' + mm[1]).slice(-2);
  if (ymd < Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')) ymd = (y + 1) + ymd.slice(4);
  return { ymd: ymd, time: ('0' + mm[3]).slice(-2) + ':00', rest: (mm[4] || '').trim() };
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
  if (!res.ok) tg_('sendMessage', { chat_id: m.chat.id, text: 'Не записав: ' + ({ full: 'на цю годину вже 2 машини', bad_plate: 'номер авто неповний — напишіть повністю або приберіть', bad_date: 'дата поза графіком (неділя або минула дата)', bad_time: 'час поза графіком (Пн–Пт 9–16, Сб 9–14)', dup: 'це авто вже записане на цей день' }[res.error] || res.error) });
}

function cmdCancel_(m, text) {
  var dt = parseDT_(text.replace(/^\/\S+\s*/, ''));
  var help = 'Формат: /cancel 28.11 10 AX1234KA\n(дата, година і номер авто або слово із запису, напр. Шкода)';
  if (!dt || !dt.rest) { tg_('sendMessage', { chat_id: m.chat.id, text: help }); return; }
  var key = normPlate_(dt.rest), word = dt.rest.toLowerCase();
  var sh = bookSheet_(), v = sh.getDataRange().getDisplayValues(), hit = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i][2] !== dt.ymd || v[i][3] !== dt.time || v[i][10] === 'скасовано') continue;
    var pl = v[i][4].replace(/\s/g, '');
    if ((key.length >= 4 && pl && pl.indexOf(key) >= 0) || v[i][5].toLowerCase().indexOf(word) >= 0) hit.push(i);
  }
  if (!hit.length) { tg_('sendMessage', { chat_id: m.chat.id, text: 'Не знайшов такого запису на ' + niceDate_(dt.ymd) + ' ' + dt.time + '. Перевірте: /plan ' + dt.ymd.slice(8) + '.' + dt.ymd.slice(5, 7) }); return; }
  if (hit.length > 1) { tg_('sendMessage', { chat_id: m.chat.id, text: 'Підходить кілька записів, уточніть:\n' + hit.map(function (i) { return '• ' + (v[i][4] || '') + ' ' + v[i][5]; }).join('\n') }); return; }
  var r = hit[0];
  sh.getRange(r + 1, 11).setValue('скасовано');
  tg_('sendMessage', { chat_id: m.chat.id, text: '❌ Скасовано: ' + niceDate_(dt.ymd) + ' ' + dt.time + ' — ' + ((v[r][4] ? v[r][4] + ' ' : '') + v[r][5]).slice(0, 150) + '\nМісце звільнилось.' });
}

function sendPlan_(chat, arg) {
  var ymd;
  if (arg && /^\d{1,2}\.\d{1,2}$/.test(arg)) {
    var p = arg.split('.'); ymd = Utilities.formatDate(new Date(), TZ, 'yyyy') + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2);
  } else {
    ymd = Utilities.formatDate(new Date(Date.now() + 864e5), TZ, 'yyyy-MM-dd');
  }
  var v = bookSheet_().getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return r[2] === ymd && r[10] !== 'скасовано'; })
    .sort(function (a, b) { return a[3] < b[3] ? -1 : 1; });
  var s = v.length ? v.map(function (r) { return r[3] + '  ' + r[4] + (r[5] ? ' · ' + r[5] : '') + '  — шини: ' + r[8] + (r[11] ? '\n        🗄 ' + r[11] : ''); }).join('\n') : 'записів немає';
  var sl = slots_(ymd);
  var free = sl.slots ? sl.slots.reduce(function (a, x) { return a + x.left; }, 0) : 0;
  tg_('sendMessage', { chat_id: chat, text: '📋 ' + niceDate_(ymd) + ' — ' + v.length + ' маш.\n' + s + '\n\nВільних місць: ' + free });
}

// Запускається тригером щодня о 19:00: план на завтра.
function dailyPlan() { if (chatId_()) sendPlan_(chatId_()); }

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
  Logger.log(JSON.stringify(slots_(Utilities.formatDate(new Date(Date.now() + 864e5), TZ, 'yyyy-MM-dd'))));
}
