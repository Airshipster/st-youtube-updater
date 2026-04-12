const SHEET_NAME   = 'Стат. Каналы';
const START_ROW    = 3;
const COL          = { LINK:1, CUSTOM:2, TITLE:3, VIDEOS:4, SUBS:5, VIEWS:6, KIDS:7, LICENSE:8, CATS:9, TAGS:10, CREATED:11, LAST:12, ID:13 };
const BATCH_SIZE   = 50;
const DATE_FMT     = 'dd.MM.yyyy';
const DAILY_HOUR   = 8; // 08:00

const RUNTIME_BUDGET_MS = 0;
const QUOTA_BUDGET      = 0;

let REQ = { channels: 0, playlistItems: 0, videos: 0 };

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('YouTube')
    .addItem('Обновить данные', 'updateYouTubeData')
    .addSeparator()
    .addItem('Включить ежедневный триггер (08:00)', 'enableDailyTrigger')
    .addItem('Отключить триггеры', 'disableAllTriggers')
    .addToUi();

  const hasRight = ScriptApp.getProjectTriggers()
    .some(t => t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK &&
               t.getHandlerFunction() === 'updateYouTubeData');
  if (!hasRight) enableDailyTrigger();
}

function getSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Лист "' + SHEET_NAME + '" не найден.');
  return sh;
}
function tz_(sheet){ try{ return sheet.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }catch(e){ return Session.getScriptTimeZone(); } }
function nowStamp_(sheet){ return Utilities.formatDate(new Date(), tz_(sheet), 'dd.MM.yyyy HH:mm:ss'); }
function clock_(sheet){ return Utilities.formatDate(new Date(), tz_(sheet), 'HH:mm:ss'); }

function logSet_(sheet, text){ sheet.getRange(2,1).setValue('[' + clock_(sheet) + '] ' + text); }
function logAdd_(sheet, text){
  const c = sheet.getRange(2,1);
  const prev = String(c.getValue() || '');
  c.setValue(prev ? prev + '\n[' + clock_(sheet) + '] ' + text : '[' + clock_(sheet) + '] ' + text);
}

function currentTriggerStatus_() {
  const clocks = ScriptApp.getProjectTriggers()
    .filter(t => t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK);
  const hasRight = clocks.some(t => t.getHandlerFunction() === 'updateYouTubeData');
  return hasRight ? 'Ежедневный триггер включён (08:00)' : 'Все триггеры отключены';
}

function setStatusInA2_(statusText){
  const sh = getSheet_();
  const cell = sh.getRange(2,1);
  let txt = String(cell.getValue() || '');

  txt = txt.replace(/\u00A0/g, ' ');
  const statuses = [
    'Ежедневный\\s+триггер\\s+включён\\s*\\(\\d{2}:\\d{2}\\)',
    'Все\\s+триггеры\\s+отключены',
    'Триггеры\\s+обновления\\s+YouTube\\s+отключены',
    'Нет\\s+таймера\\s+updateYouTubeData'
  ];
  const re = new RegExp('(?:^|\\s*\\|\\s*)(' + statuses.join('|') + ')(?=\\s*(?:\\||$))', 'g');
  txt = txt.replace(re, '');
  txt = txt.replace(/\s*\|\s*(\|\s*)+/g, ' | ')
           .replace(/^\s*\|\s*|\s*\|\s*$/g, '')
           .trim();

  const esc = statusText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const reHas = new RegExp('(?:^|\\s*\\|\\s*)' + esc + '(?=\\s*(?:\\||$))');
  if (!reHas.test(txt)) {
    txt = txt ? (txt + ' | ' + statusText) : statusText;
  }

  cell.setValue(txt);
}

function withRetry_(fn, tries){
  tries = tries || 3;
  let delay = 500;
  for (let i=0;i<tries;i++){
    try { return fn(); }
    catch(e){
      const msg = String(e && e.message || e);
      if (i < tries-1 && /(?:403|429|5\d{2})/.test(msg)) {
        Utilities.sleep(delay);
        delay = Math.min(delay*2, 4000);
      } else {
        throw e;
      }
    }
  }
}

function normalizeTopicTail_(u){
  try{
    let tail = decodeURIComponent(String(u).split('/').pop() || '').replace(/_/g,' ');
    tail = tail.replace(/^Category:\s*/i,'');
    return tail.trim();
  }catch(_){ return ''; }
}

function catMap_(){
  return {
    "Video game culture":"Культура компьютерных игр",
    "Action game":"Экшен игры",
    "Action-adventure game":"Приключенческие игры",
    "Role-playing video game":"Компьютерные ролевые игры",
    "Strategy video game":"Стратегические игры",
    "Sports game":"Спортивные симуляторы",
    "Puzzle video game":"Головоломки",
    "Racing video game":"Гоночные игры",
    "Simulation video game":"Видеоигры-симуляторы",
    "Casual game":"Казуальные игры",
    "Music video game":"Музыкальные видеоигры",
    "Music":"Музыка",
    "Pop music":"Поп-музыка",
    "Hip hop music":"Хип-хоп",
    "Electronic music":"Электронная музыка",
    "Music of Asia":"Азиатская музыка",
    "Rock music":"Рок-музыка",
    "Christian music":"Церковная музыка",
    "Music of Latin America":"Латиноамериканская музыка",
    "Independent music":"Инди-музыка",
    "Classical music":"Классическая музыка",
    "Rhythm and blues":"Ритм-энд-блюз",
    "Jazz":"Джаз",
    "Country music":"Кантри музыка",
    "Soul music":"Соул музыка",
    "Reggae":"Регги",
    "Sport":"Спорт",
    "Association football":"Футбол",
    "Motorsport":"Моторные виды спорта",
    "Physical fitness":"Фитнес",
    "American football":"Американский футбол",
    "Basketball":"Баскетбол",
    "Mixed martial arts":"Смешанные боевые искусства",
    "Baseball":"Бейсбол",
    "Boxing":"Бокс",
    "Golf":"Гольф",
    "Professional wrestling":"Реслинг",
    "Cricket":"Крикет",
    "Ice hockey":"Хоккей с шайбой",
    "Tennis":"Теннис",
    "Volleyball":"Волейбол",
    "Lifestyle (sociology)":"Образ жизни",
    "Knowledge":"Образование",
    "Technology":"Технологии",
    "Society":"Общество",
    "Entertainment":"Развлечения",
    "Film":"Кино",
    "Hobby":"Хобби",
    "Vehicle":"Транспорт",
    "Food":"Кулинария",
    "Television program":"Телевизионные программы",
    "Religion":"Религии",
    "Health":"Здоровье",
    "Politics":"Политика",
    "Tourism":"Туризм",
    "Fashion":"Мода",
    "Pet":"Домашние животные",
    "Humour":"Юмор",
    "Business":"Бизнес",
    "Performing arts":"Сценическое искусство",
    "Physical attractiveness":"Внешняя привлекательность",
    "Military":"Военное"
  };
}

function updateYouTubeData(){
  const sh = getSheet_();
  REQ = { channels: 0, playlistItems: 0, videos: 0 };

  const lastRowBefore = sh.getLastRow();
  if (lastRowBefore >= START_ROW) {
    const rows = lastRowBefore - START_ROW + 1;
    sh.getRange(START_ROW, 1, rows, 6).clearContent();
    sh.getRange(START_ROW, 9, rows, 4).clearContent();
  }

  logSet_(sh, 'Старт обновления…');

  const lastRow = sh.getLastRow();
  const n = Math.max(0, lastRow - START_ROW + 1);
  const idsAll = n ? sh.getRange(START_ROW, COL.ID, n, 1).getValues()
                     .map((r,i)=>({row:START_ROW+i, id:String(r[0]||'').trim()})) : [];
  const ids = idsAll.filter(o=>o.id && o.id.toLowerCase()!=='id');

  logAdd_(sh, 'Найдено каналов: ' + ids.length + ' | Размер батча: ' + BATCH_SIZE);
  if (ids.length === 0) {
    const endLine = '🤖 Обновление завершено: ' + nowStamp_(sh) +
                    ' | Обработано: 0 | Запросов: 0 | Квота ≈ 0 / 10000 (остаток ≈ 10000)';
    sh.getRange(2,1).setValue(endLine);
    setStatusInA2_(currentTriggerStatus_());
    return;
  }

  if (n > 0) sh.getRange(START_ROW, COL.CREATED, n, 2).setNumberFormat(DATE_FMT);

  const cmap = catMap_();
  let processed = 0;
  const t0 = Date.now();
  let spentUnits = 0;

  for (let i=0; i<ids.length; i+=BATCH_SIZE) {
    if (RUNTIME_BUDGET_MS && (Date.now() - t0 > RUNTIME_BUDGET_MS)) { logAdd_(sh, 'Достигнут бюджет времени — ранний выход'); break; }
    if (QUOTA_BUDGET && (spentUnits >= QUOTA_BUDGET))             { logAdd_(sh, 'Достигнут квотный бюджет — ранний выход'); break; }

    const batch = ids.slice(i, i+BATCH_SIZE);
    const idList = batch.map(o=>o.id).join(',');

    let chRes;
    try {
      chRes = withRetry_(() => YouTube.Channels.list(
        'snippet,statistics,contentDetails,status,brandingSettings,topicDetails', { id: idList }
      ), 3);
      REQ.channels += 1; spentUnits += 1;
    } catch (e) {
      logAdd_(sh, 'Ошибка channels.list ('+(i+1)+'-'+(i+batch.length)+'): ' + e.message);
      continue;
    }

    const dict = {};
    if (chRes && chRes.items) {
      for (const ch of chRes.items) {
        dict[ch.id] = {
          id: ch.id,
          title: ch.snippet ? ch.snippet.title : '',
          videos: ch.statistics ? ch.statistics.videoCount : '',
          subs: (ch.statistics && !ch.statistics.hiddenSubscriberCount) ? ch.statistics.subscriberCount : '(скрыты)',
          views: ch.statistics ? ch.statistics.viewCount : '',
          createdISO: ch.snippet ? ch.snippet.publishedAt : null,
          uploads: ch.contentDetails ? ch.content
