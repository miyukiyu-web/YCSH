function registerTask(type, identifier, person) {
  const number = clean_(identifier);
  const keeper = clean_(person);
  if (type !== '報廢' && type !== '轉移') throw new Error('請選擇登記類型。');
  if (type === '報廢' ? !number : !number && !keeper) {
    throw new Error(type === '報廢' ? '請填機器編號。' : '請填財產編號或新保管人。');
  }
  if (number.length > 120 || keeper.length > 120) throw new Error('登記內容超過 120 字。');
  return locked_(() => {
    const book = book_();
    const tasks = data_(book, TASKS, TASK_FIELDS.length);
    const statuses = new Map(cases_(book).rows.map(row => [row[0], row[2]]));
    const duplicate = number && tasks.rows.some(row => row[1] === type && code_(row[3]) === code_(number) &&
      code_(row[4]) === code_(keeper) && statuses.get(row[5]) !== '已辦結');
    if (duplicate) throw new Error('已有相同的登記。');
    const id = Utilities.getUuid();
    write_(book, [append_(tasks.id, [[id, type, now_(), number, keeper, '', '']])]);
    return id;
  });
}

function taskDashboard() {
  return locked_(() => {
    const book = book_();
    const cases = cases_(book).rows;
    const statuses = new Map(cases.map(row => [row[0], row[2]]));
    const active = new Map(assets_(book, ACTIVE).rows.map(row => [row[0], row]));
    const tasks = data_(book, TASKS, TASK_FIELDS.length).rows
      .filter(row => row[0] && !['待辦結', '已辦結'].includes(statuses.get(row[5])))
      .reverse().map(row => ({
        id: row[0], type: row[1], date: row[2], number: row[3], person: row[4],
        asset: active.get(row[6]) || null
      }));
    const records = cases.filter(row => row[0]).reverse().map(row => ({
      id: row[0], type: row[1], status: row[2], date: row[3], closed: row[4],
      items: JSON.parse(row[5]).map(item => item.asset[0])
    }));
    return {tasks, ready: records.filter(record => record.status === '待辦結'),
      history: records.filter(record => record.status !== '待辦結')};
  });
}

function prepareTask(id, selectedCode) {
  return locked_(() => {
    const book = book_();
    const {tasks, index, cases} = openTask_(book, id);
    const task = tasks.rows[index];
    const active = assets_(book, ACTIVE).rows;
    const prepared = task[6] && active.find(row => row[0] === task[6]);
    if (prepared) return {asset: prepared};
    const key = code_(task[3]);
    const exact = key ? active.filter(row => code_(row[0]) === key) : [];
    let matches = exact.length ? exact : active.filter(row => key && code_(row[0]).split(' ').includes(key));
    if (matches.length > 1 && task[1] === '報廢' && task[4]) {
      const keeper = code_(task[4]);
      matches = matches.filter(row => code_(row[16]) === keeper);
    }
    if (selectedCode) matches = active.filter(row => code_(row[0]) === code_(selectedCode));
    if (matches.length !== 1) return {matches};
    const asset = matches[0];
    const statuses = new Map(cases.map(row => [row[0], row[2]]));
    const inOtherTask = tasks.rows.some(row => row[0] !== id && row[6] === asset[0] &&
      (!row[5] || statuses.get(row[5]) === '已取消'));
    const inCase = cases.some(row => row[2] === '待辦結' &&
      JSON.parse(row[5]).some(item => item.asset[0] === asset[0]));
    if (inOtherTask || inCase) throw new Error('這筆財產已在其他已準備清單。');
    write_(book, [{updateCells: {
      start: {sheetId: tasks.id, rowIndex: index + 1, columnIndex: 6},
      rows: [cells_([asset[0]])], fields: 'userEnteredValue'
    }}]);
    return {asset};
  });
}

function cancelTask(id) {
  return locked_(() => {
    const book = book_();
    const {tasks, index} = openTask_(book, id);
    const task = tasks.rows[index];
    if (task[6] && assets_(book, ACTIVE).rows.some(row => row[0] === task[6])) {
      throw new Error('這筆資料已移至已準備清單，請重新整理。');
    }
    write_(book, [{deleteDimension: {range: {
      sheetId: tasks.id, dimension: 'ROWS', startIndex: index + 1, endIndex: index + 2
    }}}]);
  });
}

function openTask_(book, id) {
  const tasks = data_(book, TASKS, TASK_FIELDS.length);
  const index = tasks.rows.findIndex(row => row[0] === id);
  if (index < 0) throw new Error('找不到這筆登記，請重新整理。');
  const cases = cases_(book).rows;
  if (tasks.rows[index][5] && !cases.some(row => row[0] === tasks.rows[index][5] && row[2] === '已取消')) {
    throw new Error('這筆登記已建立申請單，請重新整理。');
  }
  return {tasks, index, cases};
}

function processPendingTasks(action, ids) {
  if (!['prepare', 'abandon'].includes(action) || !Array.isArray(ids) || !ids.length ||
      ids.some(id => typeof id !== 'string' || !id)) {
    throw new Error('請先勾選待辦，再選擇批次操作。');
  }
  return Array.from(new Set(ids), id => {
    try {
      if (action === 'abandon') {
        cancelTask(id);
        return {id, done: true};
      }
      const result = prepareTask(id);
      return result.asset ? {id, done: true} : {id, matches: result.matches};
    } catch (error) {
      return {id, error: error.message || String(error)};
    }
  });
}
