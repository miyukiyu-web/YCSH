function unzip_(bytes) {
  return Utilities.unzip(Utilities.newBlob(bytes, 'application/zip'));
}

function xlsxRows_(bytes) {
  const files = new Map(unzip_(bytes).map(blob => [blob.getName(), blob]));
  function root(path) {
    const blob = files.get(path);
    if (!blob) throw new Error('Excel 缺少必要資料：' + path);
    return XmlService.parse(blob.getDataAsString('UTF-8')).getRootElement();
  }
  function child(node, name) { return node.getChild(name, node.getNamespace()); }
  function children(node, name) { return node.getChildren(name, node.getNamespace()); }
  function attr(node, name) { return node.getAttribute(name).getValue(); }
  function cellText(node) {
    const direct = child(node, 't');
    return direct ? direct.getText() : children(node, 'r').map(run => {
      const t = child(run, 't');
      return t ? t.getText() : '';
    }).join('');
  }
  const workbook = root('xl/workbook.xml');
  const sheet = children(child(workbook, 'sheets'), 'sheet').find(node => attr(node, 'name') === 'Sheet1');
  if (!sheet) throw new Error('找不到 Sheet1。');
  const relNs = XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  const relId = sheet.getAttribute('id', relNs).getValue();
  const rels = root('xl/_rels/workbook.xml.rels');
  const relation = children(rels, 'Relationship').find(node => attr(node, 'Id') === relId);
  if (!relation) throw new Error('無法定位 Sheet1。');
  const target = attr(relation, 'Target');
  const xmlPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
  const strings = files.has('xl/sharedStrings.xml') ?
    children(root('xl/sharedStrings.xml'), 'si').map(cellText) : [];
  const rows = [];
  children(child(root(xmlPath), 'sheetData'), 'row').forEach(row => {
    const values = Array(13).fill('');
    children(row, 'c').forEach(cell => {
      const column = attr(cell, 'r').match(/^[A-Z]+/)[0];
      if (column.length !== 1 || column > 'M') return;
      const value = child(cell, 'v');
      const type = cell.getAttribute('t');
      const kind = type ? type.getValue() : '';
      if (kind === 'inlineStr') values[column.charCodeAt(0) - 65] = cellText(child(cell, 'is'));
      else if (value) values[column.charCodeAt(0) - 65] = kind === 's' ? strings[Number(value.getText())] : value.getText();
    });
    rows[Number(attr(row, 'r')) - 1] = values;
  });
  return rows;
}

function importAssets(name, encoded, commit, assetName) {
  if (!/\.xlsx$/i.test(name) || typeof encoded !== 'string' || encoded.length > 7000000) {
    throw new Error('請選擇 5 MB 以下的 .xlsx 清冊。');
  }
  const filter = clean_(assetName).normalize('NFKC') || '個人電腦';
  const parsed = parseInventory_(xlsxRows_(Utilities.base64Decode(encoded)), name, filter);
  return locked_(() => {
    const book = book_();
    const active = assets_(book, ACTIVE);
    const live = new Set(active.rows.map(row => row[0]));
    const archive = new Set([SCRAP, MOVE].flatMap(sheet => assets_(book, sheet).rows.map(row => row[0])));
    let duplicates = 0;
    let archived = 0;
    const fresh = parsed.rows.filter(row => {
      if (archive.has(row[0])) { archived++; return false; }
      if (live.has(row[0])) { duplicates++; return false; }
      live.add(row[0]);
      return true;
    });
    if (commit && parsed.errors.length) throw new Error('請修正預覽中的異常資料後再匯入。');
    if (commit && fresh.length) write_(book, [append_(active.id, fresh)]);
    return {matched: parsed.matched, skipped: parsed.skipped, fresh: fresh.length, duplicates, archived, errors: parsed.errors, written: commit ? fresh.length : 0};
  });
}

function parseInventory_(rows, name, assetName) {
  const school = rows.slice(0, 3).some(row => row.some(value => clean_(value).includes('永慶高級中學')));
  const header = rows.findIndex(row => /財產編號\s*登錄號/.test(row[0]) && row[1] === '財產名稱' &&
    /保管人員/.test(row[10]));
  if (!school || header < 0) throw new Error('Sheet1 不是永慶高中財產盤點清冊。');
  const result = {rows: [], errors: [], matched: 0, skipped: 0};
  const split = value => clean_(value).split(/\r?\n/).map(clean_);
  const footer = rows.findIndex((row, index) => index > header &&
    /^(說明[:：]|保管人\s+財產管理單位)/.test(clean_(row[0])));
  rows.slice(header + 1, footer < 0 ? rows.length : footer).forEach((row, index) => {
    const code = code_(row[0]);
    const label = clean_(row[1]).normalize('NFKC');
    if (!label || (!code && !clean_(row[4]))) return;
    if (label !== assetName) {
      result.skipped++;
      return;
    }
    result.matched++;
    const qty = Number(row[4]);
    if (!/^[0-9][A-Z0-9-]* \d+$/.test(code) || !Number.isInteger(qty) || qty < 1) {
      result.errors.push('第 ' + (header + index + 2) + ' 列：編號或數量有誤');
      return;
    }
    const [model, extra = ''] = split(row[2]);
    const [residual, depreciation = ''] = split(row[5]);
    const [original, bookValue = ''] = split(row[6]);
    const [purchased, issued = ''] = split(row[7]);
    const [life, elapsed = ''] = split(row[8]);
    const owners = split(row[10]);
    result.rows.push([code, clean_(row[1]), model, extra === model ? '' : extra, clean_(row[3]), qty,
      residual, depreciation, original, bookValue, purchased, issued, life, elapsed,
      clean_(row[9]), owners.length > 1 ? owners[0] : '', owners.length > 1 ? owners.slice(1).join(' ') : owners[0],
      clean_(row[11]), clean_(row[12]), name]);
  });
  return result;
}
