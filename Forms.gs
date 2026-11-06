const WORD_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function usedAge_(purchase, date) {
  const match = clean_(purchase).match(/^(\d{1,3}|\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!match) return '';
  const startYear = Number(match[1]) + (match[1].length < 4 ? 1911 : 0);
  const startMonth = Number(match[2]);
  const startDay = Number(match[3]);
  const start = new Date(Date.UTC(startYear, startMonth - 1, startDay));
  if (start.getUTCMonth() !== startMonth - 1 || start.getUTCDate() !== startDay) return '';
  const [year, month, day] = date.split('-').map(Number);
  const months = (year - startYear) * 12 + month - startMonth - (day < startDay ? 1 : 0);
  return months < 0 ? '' : Math.floor(months / 12) + '年' + months % 12 + '月';
}

function getPrintForm(id) {
  const row = locked_(() => cases_(book_()).rows.find(record => record[0] === id));
  if (!row) throw new Error('找不到申請單。');
  if (row[2] === '已取消') throw new Error('此單已取消。');
  const move = row[1] === '轉移';
  const formDate = move ? row[3] : today_();
  const items = JSON.parse(row[5]);
  const values = items.map(item => {
    const asset = item.asset;
    const date = asset[10].replace(/[/-]/g, '');
    const shared = [date, asset[0], asset[1], asset[2], asset[4], asset[5]];
    return shared.concat(move ? [asset[16], asset[14], item.unit, item.keeper, item.location] :
      [item.price, item.total, asset[12], usedAge_(asset[10], formDate), item.reason, asset[16]]);
  });
  const columns = ['購置日期', '財產編號', '品名', '廠牌型號', '單位', '數量'].concat(move ?
    ['移出人', '移出地點', '移入單位', '移入人員', '移入地點'] :
    ['單價', '總價', '使用年限', '已使用年月', '減損報廢原因', '保管人']);
  const signatures = move ? ['原保管人', '原單位主管', '移入保管人', '移入單位主管', '庶務組'] :
    ['保管人', '主任', '會計', '校長', '會庶務組', '查簽'];
  return {move, formDate, values, columns, signatures};
}

function downloadForm(id) {
  const {move, formDate, values} = getPrintForm(id);
  const parts = unzip_(Utilities.base64Decode(move ? TRANSFER_TEMPLATE : SCRAP_TEMPLATE));
  let found = false;
  const output = parts.map(part => {
    const path = part.getName();
    if (path === 'word/document.xml') {
      found = true;
      let xml = fillWordTables_(part.getDataAsString('UTF-8'), values, move);
      if (move) xml = replaceWordDate_(xml, formDate);
      return Utilities.newBlob(xml, 'application/xml', path);
    }
    if (!move && /^word\/header\d+\.xml$/.test(path)) {
      return Utilities.newBlob(replaceWordDate_(part.getDataAsString('UTF-8'), formDate), 'application/xml', path);
    }
    return part;
  });
  if (!found) throw new Error('Word 範本缺少文件內容。');
  const name = (move ? '財產移動申請單_' : '減損報銷單_') + id + '.docx';
  return {name, mime: WORD_MIME, data: Utilities.base64Encode(Utilities.zip(output, name).getBytes())};
}

function xmlText_(value) {
  return String(value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function fillWordTables_(xml, values, move) {
  let copies = 0;
  const columns = move ? 11 : 12;
  const capacity = move ? 4 : 22;
  const output = xml.replace(/<w:tbl(?=[ >])[\s\S]*?<\/w:tbl>/g, table => {
    copies++;
    let row = -1;
    const result = table.replace(/<w:tr(?=[ >])[\s\S]*?<\/w:tr>/g, line => {
      if (++row === 0) return line;
      let column = 0;
      const updated = line.replace(/<w:tc(?=[ >])[\s\S]*?<\/w:tc>/g, cell => {
        const value = (values[row - 1] || [])[column++];
        const empty = cell.replace(/(<w:t(?:\s[^>]*)?>)[\s\S]*?<\/w:t>/g, '$1</w:t>');
        if (value === undefined || value === '') return empty;
        return empty.replace('</w:p>', () => '<w:r><w:t xml:space="preserve">' + xmlText_(value) + '</w:t></w:r></w:p>');
      });
      if (column !== columns) throw new Error('Word 範本欄位數不符。');
      return updated;
    });
    if (row !== capacity) throw new Error('Word 範本列數不符。');
    return result;
  });
  if (copies !== 2) throw new Error('Word 範本缺少第二份表格。');
  return output;
}

function replaceWordDate_(xml, date) {
  const [year, month, day] = date.split('-').map(Number);
  return xml.replace(/<w:p(?=[ >])[\s\S]*?<\/w:p>/g, paragraph => {
    const tag = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g;
    const text = Array.from(paragraph.matchAll(tag), m => m[2]).join('');
    const match = text.match(/\d{3}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*$/);
    if (!match) return paragraph;
    const replacements = Array.from(match[0].matchAll(/\d+/g), (m, i) => ({
      start: match.index + m.index, end: match.index + m.index + m[0].length,
      value: String([year - 1911, month, day][i])
    }));
    let offset = 0;
    return paragraph.replace(tag, (_, open, original, close) => {
      const start = offset;
      offset += original.length;
      let value = original;
      for (const edit of replacements.slice().reverse()) {
        if (edit.start >= offset || edit.end <= start) continue;
        value = value.slice(0, Math.max(0, edit.start - start)) +
          (edit.start >= start ? edit.value : '') + value.slice(Math.min(original.length, edit.end - start));
      }
      return open + value + close;
    });
  });
}
