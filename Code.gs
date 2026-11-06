const ACTIVE = '現行財產清冊';
const SCRAP = '報廢紀錄';
const MOVE = '轉移紀錄';
const CASES = '申請單';
const TASKS = '待辦登記';
const TASK_FIELDS = ['登記編號','類型','登記時間','機器或財產編號','保管人或新保管人','申請單號','已比對財產編號'];
const ASSET_FIELDS = [
  '完整財產編號','財產名稱','廠牌型號','品名補充','單位','數量','殘值','累計折舊',
  '原始金額','帳面價值','購置日期','領用日期','使用年限','盤點已用年限',
  '存放地點','保管代碼','保管人','備註','檢查情形','匯入來源'
];
const ARCHIVE_FIELDS = ['單號','辦結時間','報廢原因','移入單位','新保管人','新地點','列入現行清冊'];
const CASE_FIELDS = ['單號','類型','狀態','填單日期','辦結時間','申請明細 JSON'];
const DEFAULT_REASON = '年限已到不堪使用';

function setup() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('請在系統專用 Google 試算表的繫結專案執行 setup。');
  return locked_(() => {
    const layouts = {
      [ACTIVE]: ASSET_FIELDS,
      [SCRAP]: ASSET_FIELDS.concat(ARCHIVE_FIELDS),
      [MOVE]: ASSET_FIELDS.concat(ARCHIVE_FIELDS),
      [CASES]: CASE_FIELDS,
      [TASKS]: TASK_FIELDS
    };
    for (const [name, headers] of Object.entries(layouts)) {
      const sheet = book.getSheetByName(name) || book.insertSheet(name);
      if (sheet.getMaxColumns() < headers.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
      }
      if (!sheet.getLastRow()) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sheet.setFrozenRows(1);
      } else {
        const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
        if (name === TASKS && actual.slice(0, -1).join('|') === headers.slice(0, -1).join('|') &&
            !actual[6] && (sheet.getLastRow() === 1 ||
              sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).isBlank())) {
          sheet.getRange(1, 7).setValue(headers[6]).setFontWeight('bold');
        } else if (actual.join('|') !== headers.join('|')) {
          throw new Error('工作表欄位不符：' + name + '。請檢查表頭。');
        }
      }
    }
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty('ASSET_BOOK_ID', book.getId());
    return book.getUrl();
  });
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('永慶高中財產報廢與轉移');
}

function locked_(action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return action(); } finally { lock.releaseLock(); }
}

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('ASSET_BOOK_ID');
  if (!id) throw new Error('請先在試算表的 Apps Script 執行 setup。');
  return SpreadsheetApp.openById(id);
}

function data_(book, name, width) {
  const sheet = book.getSheetByName(name);
  const last = sheet.getLastRow();
  const rows = last > 1 ? sheet.getRange(2, 1, last - 1, width).getValues() : [];
  return {id: sheet.getSheetId(), rows};
}

function assets_(book, name) {
  return data_(book, name, ASSET_FIELDS.length + (name === ACTIVE ? 0 : ARCHIVE_FIELDS.length));
}

function cases_(book) {
  return data_(book, CASES, CASE_FIELDS.length);
}

function cells_(values) {
  return {values: values.map(value => ({
    userEnteredValue: typeof value === 'number' ? {numberValue: value} : {stringValue: String(value)}
  }))};
}

function append_(sheetId, rows) {
  return {appendCells: {sheetId, rows: rows.map(cells_), fields: 'userEnteredValue'}};
}

function write_(book, requests) {
  if (requests.length) Sheets.Spreadsheets.batchUpdate({requests}, book.getId());
}

function clean_(value) {
  return String(value == null ? '' : value).trim();
}

function code_(value) {
  return clean_(value).normalize('NFKC').replace(/\s+/g, ' ').toUpperCase();
}

function today_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
}

function searchAssets(query, page) {
  const key = code_(query);
  if (!key) throw new Error('請輸入財產編號、登錄號或保管人。');
  return locked_(() => {
    const book = book_();
    const pending = new Set(cases_(book).rows.filter(row => row[2] === '待辦結')
      .flatMap(row => JSON.parse(row[5]).map(item => item.asset[0])));
    const exact = /^[A-Z0-9-]+ \d+$/.test(key);
    const hits = [ACTIVE, SCRAP, MOVE].flatMap(name =>
      assets_(book, name).rows.filter(row => exact ? row[0] === key :
        [row[0], row[16], row[24]].some(value => code_(value).includes(key)))
        .map(row => ({kind: name, row, pending: name === ACTIVE && pending.has(row[0])}))
    );
    const offset = Math.max(0, Number(page) || 0) * 25;
    return {total: hits.length, hits: hits.slice(offset, offset + 25)};
  });
}

function listCases(type) {
  return locked_(() => cases_(book_()).rows.filter(row => row[1] === type).reverse()
    .map(row => ({id: row[0], status: row[2], date: row[3], closed: row[4],
      items: JSON.parse(row[5]).map(item => item.asset[0])})));
}

function createCase(payload) {
  return locked_(() => {
    const {type, items} = payload;
    const limit = type === '報廢' ? 22 : type === '轉移' ? 4 : 0;
    if (!Array.isArray(items) || !items.length || items.length > limit) {
      throw new Error('所選筆數超出表單可用列數。');
    }
    if (type === '轉移' && items.length > 1 && items.some(item => item.taskId)) {
      throw new Error('待辦登記請逐筆建立申請單，以便分別完成並清除。');
    }
    const book = book_();
    const cases = cases_(book);
    const active = new Map(assets_(book, ACTIVE).rows.map(row => [row[0], row]));
    const pending = new Set(cases.rows.filter(row => row[2] === '待辦結')
      .flatMap(row => JSON.parse(row[5]).map(item => item.asset[0])));
    const tasks = data_(book, TASKS, TASK_FIELDS.length);
    const caseStatus = new Map(cases.rows.map(row => [row[0], row[2]]));
    const taskIds = new Set();
    const taskUpdates = [];
    const prepared = new Map(tasks.rows.filter(row => row[6] &&
      (!row[5] || caseStatus.get(row[5]) === '已取消')).map(row => [row[6], row[0]]));
    const detail = items.map(item => {
      const key = code_(item.code);
      if (prepared.has(key) && prepared.get(key) !== item.taskId) {
        throw new Error('財產已在其他待辦的已準備清單：' + key);
      }
      if (item.taskId) {
        const index = tasks.rows.findIndex(row => row[0] === item.taskId);
        const task = tasks.rows[index];
        if (!task || task[1] !== type || task[6] !== key || taskIds.has(item.taskId) ||
            (task[5] && caseStatus.get(task[5]) !== '已取消')) {
          throw new Error('此登記已送單或無法使用，請重新整理待辦。');
        }
        taskIds.add(item.taskId);
        taskUpdates.push(index + 1);
      }
      const asset = active.get(key);
      if (!asset || pending.has(key)) throw new Error('財產無法送單：' + key);
      pending.add(key);
      if (type === '報廢') {
        const reason = clean_(item.reason) || DEFAULT_REASON;
        const total = clean_(item.total);
        const price = clean_(item.price);
        if (!/^\d+(?:\.\d{1,2})?$/.test(price) || !/^\d+(?:\.\d{1,2})?$/.test(total) ||
            Math.abs(Number(price) * Number(asset[5]) - Number(total)) > 0.005) {
          throw new Error(key + ' 的單價、數量與總價須一致。');
        }
        if (!usedAge_(asset[10], today_())) throw new Error(key + ' 的購置日期有誤，請先確認清冊。');
        return {asset, reason, total, price};
      }
      const unit = clean_(item.unit);
      const keeper = clean_(item.keeper);
      const location = clean_(item.location);
      if (!unit || !keeper || !location) throw new Error(key + ' 請填完整移入資料。');
      return {asset, unit, keeper, location, keepActive: item.keepActive === true};
    });
    const date = today_();
    const id = (type === '報廢' ? 'S-' : 'T-') + date.replace(/-/g, '') + '-' + Utilities.getUuid();
    const json = JSON.stringify(detail);
    if (json.length > 45000) throw new Error('明細過長，請分單。');
    const requests = [append_(cases.id, [[id, type, '待辦結', date, '', json]])];
    taskUpdates.forEach(rowIndex => requests.push({updateCells: {
      start: {sheetId: tasks.id, rowIndex, columnIndex: 5},
      rows: [cells_([id])], fields: 'userEnteredValue'
    }}));
    write_(book, requests);
    return id;
  });
}

function finishCase(id, action) {
  if (!['close','cancel'].includes(action)) throw new Error('未知作業。');
  return locked_(() => {
    const book = book_();
    const cases = cases_(book);
    const index = cases.rows.findIndex(row => row[0] === id);
    if (index < 0) throw new Error('找不到申請單。');
    const record = cases.rows[index];
    const status = action === 'close' ? '已辦結' : '已取消';
    if (record[2] === status) return status;
    if (record[2] !== '待辦結') throw new Error('這張單已處理。');
    const time = now_();
    const requests = [];
    if (action === 'close') {
      const isMove = record[1] === '轉移';
      const history = assets_(book, isMove ? MOVE : SCRAP);
      const active = assets_(book, ACTIVE);
      const detail = JSON.parse(record[5]);
      const positions = new Map(active.rows.map((row, i) => [row[0], i + 1]));
      const remove = [];
      const receive = [];
      const archive = detail.map(item => {
        const key = item.asset[0];
        const at = positions.get(key);
        if (at === undefined || JSON.stringify(active.rows[at - 1]) !== JSON.stringify(item.asset)) {
          throw new Error('現行清冊已變更，請取消此單後重新查詢：' + key);
        }
        remove.push(at);
        if (isMove && item.keepActive) {
          const updated = item.asset.slice();
          updated[14] = item.location;
          updated[15] = '';
          updated[16] = item.keeper;
          receive.push(updated);
        }
        return item.asset.concat([id, time, item.reason || '', item.unit || '',
          item.keeper || '', item.location || '', item.keepActive ? '是' : '否']);
      });
      requests.push(append_(history.id, archive));
      remove.sort((a, b) => b - a).forEach(at => requests.push({
        deleteDimension: {range: {sheetId: active.id, dimension: 'ROWS', startIndex: at, endIndex: at + 1}}
      }));
      if (receive.length) requests.push(append_(active.id, receive));
    }
    requests.push({updateCells: {
      start: {sheetId: cases.id, rowIndex: index + 1, columnIndex: 2},
      rows: [cells_([status, record[3], time])], fields: 'userEnteredValue'
    }});
    write_(book, requests);
    return status;
  });
}

function clearAllData(confirmation) {
  if (confirmation !== '清除全部資料') throw new Error('確認文字不符。');
  return locked_(() => {
    const book = book_();
    const layouts = [
      [ACTIVE, ASSET_FIELDS.length],
      [SCRAP, ASSET_FIELDS.length + ARCHIVE_FIELDS.length],
      [MOVE, ASSET_FIELDS.length + ARCHIVE_FIELDS.length],
      [CASES, CASE_FIELDS.length],
      [TASKS, TASK_FIELDS.length]
    ];
    const requests = [];
    const counts = {};
    for (const [name, width] of layouts) {
      const sheet = book.getSheetByName(name);
      if (!sheet) throw new Error('缺少「' + name + '」工作表，請先執行 setup。');
      const last = sheet.getLastRow();
      counts[name] = Math.max(0, last - 1);
      if (last <= 1) continue;
      requests.push({updateCells: {
        range: {sheetId: sheet.getSheetId(), startRowIndex: 1, endRowIndex: last,
          startColumnIndex: 0, endColumnIndex: width},
        rows: [], fields: 'userEnteredValue'
      }});
    }
    write_(book, requests);
    return counts;
  });
}
