const EXPENSE = {
  timeZone: 'Asia/Tokyo',
  sheet: '経費明細',
  fares: '運賃マスター',
  maxItems: 8,
  photosPerItem: 5,
  maxPhotos: 12,
  maxImage: 900 * 1024,
  headers: ['申請日', '受付番号', '明細番号', '氏名', '利用日', '区分', '用途', '申請額（円）',
    '自宅最寄駅', '事務所最寄駅', '集合駐輪場最寄駅', '運賃種別', '対象区間', '集合駅まで（円）',
    '事務所まで（円）', '集合駅までの経路', '事務所までの経路', '運賃の入力方法',
    '写真1', '写真2', '写真3', '写真4', '写真5', '画像の読み取り結果', '受付日時（JST）', '確認状況'],
  fareHeaders: ['出発駅', '到着駅', '運賃種別（IC / 切符）', '対象区間（片道 / 往復）', '運賃（円）',
    '経路・鉄道会社', '適用開始日', '適用終了日'],
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('経費申請')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function setupExpenseApp() {
  const activeEmail = Session.getActiveUser().getEmail();
  if (!activeEmail || activeEmail !== Session.getEffectiveUser().getEmail()) throw new Error('管理者用の操作です。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    let folder = props.getProperty('EXPENSE_FOLDER_ID');
    if (!folder) {
      folder = Drive.Files.create({ name: '経費申請｜証憑', mimeType: 'application/vnd.google-apps.folder' }, null, { fields: 'id' }).id;
      props.setProperty('EXPENSE_FOLDER_ID', folder);
    }
    let book = props.getProperty('EXPENSE_BOOK_ID');
    if (!book) {
      const created = Sheets.Spreadsheets.create({
        properties: { title: '経費申請台帳', timeZone: EXPENSE.timeZone, locale: 'ja_JP' },
        sheets: [
          { properties: { sheetId: 101, title: EXPENSE.sheet, gridProperties: { frozenRowCount: 1, columnCount: 26 } } },
          { properties: { sheetId: 102, title: EXPENSE.fares, gridProperties: { frozenRowCount: 1 }} },
        ],
      });
      book = created.spreadsheetId;
      props.setProperty('EXPENSE_BOOK_ID', book);
    }
    const settings = Sheets.Spreadsheets.get(book, { fields: 'spreadsheetUrl,sheets(properties(sheetId,title))' });
    const tabs = Object.fromEntries(settings.sheets.map(({ properties }) => [properties.title, properties.sheetId]));
    if (!(EXPENSE.sheet in tabs) || !(EXPENSE.fares in tabs)) throw new Error('経費明細と運賃マスターのタブが必要です。');
    const request = { valueInputOption: 'RAW', data: [
      { range: "'" + EXPENSE.sheet + "'!A1:Z1", values: [EXPENSE.headers] },
      { range: "'" + EXPENSE.fares + "'!A1:H1", values: [EXPENSE.fareHeaders] },
    ] };
    Sheets.Spreadsheets.Values.batchUpdate(request, book);
    const formatting = [
      { sheetId: tabs[EXPENSE.sheet], widths: [112, 305, 85, 142, 112, 160, 280, 120, 120, 120, 120, 100, 100, 112, 112, 195, 195, 205, 138, 138, 138, 138, 138, 350, 180, 110] },
      { sheetId: tabs[EXPENSE.fares], widths: [155, 155, 165, 186, 110, 260, 135, 135] },
    ].flatMap(({ sheetId, widths }) => [
      { repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: widths.length },
        cell: { userEnteredFormat: { backgroundColor: { red: .08, green: .19, blue: .29 }, textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      } },
      ...widths.map((pixelSize, index) => ({ updateDimensionProperties: {
        range: { sheetId, dimension: 'COLUMNS', startIndex: index, endIndex: index + 1 },
        properties: { pixelSize }, fields: 'pixelSize',
      } })),
    ]);
    // Keep the old column positions and historical values in existing ledgers.
    formatting.push({ updateDimensionProperties: {
      range: { sheetId: tabs[EXPENSE.sheet], dimension: 'COLUMNS', startIndex: 6, endIndex: 7 },
      properties: { hiddenByUser: true }, fields: 'hiddenByUser',
    } });
    Sheets.Spreadsheets.batchUpdate({ requests: formatting }, book);
    const folderData = Drive.Files.get(folder, { fields: 'webViewLink' });
    console.log(JSON.stringify({ 台帳: settings.spreadsheetUrl, 証憑フォルダ: folderData.webViewLink || 'https://drive.google.com/drive/folders/' + folder }));
  } finally {
    lock.releaseLock();
  }
}

function getExpenseSettings() {
  const book = PropertiesService.getScriptProperties().getProperty('EXPENSE_BOOK_ID');
  const folder = PropertiesService.getScriptProperties().getProperty('EXPENSE_FOLDER_ID');
  return { ready: Boolean(book && folder), today: Utilities.formatDate(new Date(), EXPENSE.timeZone, 'yyyy-MM-dd') };
}

function lookupExpenseFares(input) {
  const value = fareInputs_(input, true);
  const storage = expenseStorage_();
  const rows = fareRows_(storage.book);
  return {
    destination: value.destination ? fareOptions_(rows, value.home, value.destination, value) : [],
    office: value.office ? fareOptions_(rows, value.home, value.office, value) : [],
  };
}

function readExpensePhoto(image) {
  expenseStorage_();
  const blob = expensePhoto_(image, 'receipt.jpg');
  const doc = Drive.Files.create({ name: '経費申請 OCR 一時ファイル', mimeType: 'application/vnd.google-apps.document' },
    blob, { ocrLanguage: 'ja', fields: 'id' });
  try {
    const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(doc.id) + '/export?mimeType=text%2Fplain';
    const result = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    return result.getContentText('UTF-8').trim().slice(0, 1600);
  } finally {
    try {
      Drive.Files.update({ trashed: true }, doc.id, null, { fields: 'id' });
    } catch (error) {
      console.error('OCRの一時ファイルをゴミ箱に移動できませんでした。');
    }
  }
}

function submitExpense(data) {
  if (!data || typeof data.id !== 'string' || !/^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(data.id)) {
    throw new Error('受付番号を確認してください。');
  }
  const name = textValue_(data.name, 80, '氏名');
  if (!Array.isArray(data.items) || !data.items.length || data.items.length > EXPENSE.maxItems) throw new Error('明細は1〜8件で入力してください。');
  const storage = expenseStorage_();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ids = Sheets.Spreadsheets.Values.get(storage.book, "'" + EXPENSE.sheet + "'!B2:B").values || [];
    if (ids.some((row) => row[0] === data.id)) return { id: data.id, count: ids.filter((row) => row[0] === data.id).length };
    const receivedAt = new Date();
    const today = Utilities.formatDate(receivedAt, EXPENSE.timeZone, 'yyyy-MM-dd');
    const master = data.items.some((item) => item?.category === 'train' && (item.fare?.destinationKey || item.fare?.officeKey)) ? fareRows_(storage.book) : [];
    const items = data.items.map((item, index) => parseExpense_(item, index, today, master));
    if (items.reduce((total, item) => total + item.photos.length, 0) > EXPENSE.maxPhotos) throw new Error('写真は申請あたり12枚までです。');
    const now = Utilities.formatDate(receivedAt, EXPENSE.timeZone, 'yyyy-MM-dd HH:mm:ss');
    const existing = Drive.Files.list({
      q: "'" + storage.folder + "' in parents and trashed = false and name contains '" + data.id + "_'",
      pageSize: 100, fields: 'files(id,name,webViewLink)',
    }).files || [];
    const files = new Map(existing.map((file) => [file.name, file]));
    const alreadyStored = new Set();
    const rows = items.map((item, index) => {
      const links = Array(EXPENSE.photosPerItem).fill('');
      item.photos.forEach((photo, photoIndex) => {
        const hash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, photo.blob.getBytes())).replace(/=+$/, '');
        const filename = data.id + '_' + (index + 1) + '_' + (photoIndex + 1) + '_' + hash + '.jpg';
        const file = files.get(filename) || Drive.Files.create({
          name: filename, parents: [storage.folder], mimeType: 'image/jpeg',
        }, photo.blob, { fields: 'id,webViewLink' });
        alreadyStored.add(file.id);
        links[photoIndex] = file.webViewLink || 'https://drive.google.com/file/d/' + file.id + '/view';
      });
      return [today, data.id, index + 1, name, item.date, item.category, '', item.amount,
        ...item.fare, ...links,
        item.photos.map((p, i) => p.ocr ? '写真' + (i + 1) + '\n' + p.ocr : '').filter(Boolean).join('\n\n'),
        now, '未確認'];
    });
    Sheets.Spreadsheets.Values.append({ values: rows }, storage.book, "'" + EXPENSE.sheet + "'!A:Z", {
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    });
    for (const file of existing) {
      if (alreadyStored.has(file.id)) continue;
      try { Drive.Files.update({ trashed: true }, file.id, null, { fields: 'id' }); }
      catch (error) { console.error('重複写真を整理できませんでした。'); }
    }
    return { id: data.id, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function parseExpense_(item, index, today, master) {
  const categories = { general: '一般経費', train: 'キャラバン交通費', bike: '自転車引き出し代' };
  if (!item || !Object.prototype.hasOwnProperty.call(categories, item.category)) throw new Error('経費の区分を確認してください。');
  const date = dateValue_(item.date);
  if (!date || date > today) throw new Error('明細' + (index + 1) + 'の利用日を確認してください。');
  if (!Array.isArray(item.photos) || !item.photos.length || item.photos.length > EXPENSE.photosPerItem) {
    throw new Error('明細' + (index + 1) + 'に写真を1〜5枚添付してください。');
  }
  const photos = item.photos.map((photo) => ({
    blob: expensePhoto_(photo?.image, 'receipt.jpg'), ocr: typeof photo?.ocr === 'string' ? photo.ocr.trim().slice(0, 1600) : '',
  }));
  let amount;
  let fare = Array(10).fill('');
  if (item.category === 'train') {
    const trip = fareInputs_({ ...item.fare, date });
    const destination = readChosenFare_(master, trip.home, trip.destination, trip, item.fare.destinationKey, item.fare.destinationAmount);
    const office = readChosenFare_(master, trip.home, trip.office, trip, item.fare.officeKey, item.fare.officeAmount);
    amount = Math.max(0, destination.amount - office.amount);
    fare = [trip.home, trip.office, trip.destination, trip.method, trip.direction, destination.amount, office.amount,
      destination.route, office.route, '集合駅: ' + destination.source + ' / 事務所: ' + office.source];
  } else {
    amount = yen_(item.amount);
  }
  return { date, category: categories[item.category], amount, fare, photos };
}

function fareInputs_(input, allowPartial = false) {
  if (!input || !['IC', '切符'].includes(input.method) || !['片道', '往復'].includes(input.direction) || !dateValue_(input.date)) {
    throw new Error('利用日・運賃種別・片道／往復を確認してください。');
  }
  return { date: dateValue_(input.date), method: input.method, direction: input.direction,
    home: textValue_(input.home, 55, '自宅最寄駅'),
    destination: allowPartial && input.destination === '' ? '' : textValue_(input.destination, 55, '集合駐輪場最寄駅'),
    office: allowPartial && input.office === '' ? '' : textValue_(input.office, 55, '事務所最寄駅') };
}

function fareRows_(book) {
  return Sheets.Spreadsheets.Values.get(book, "'" + EXPENSE.fares + "'!A2:H", {
    valueRenderOption: 'FORMATTED_VALUE',
  }).values || [];
}

function station_(value) {
  return String(value ?? '').normalize('NFKC').replace(/\s+/g, '').replace(/駅$/, '').toLowerCase();
}

function fareOptions_(rows, from, to, trip) {
  const start = station_(from);
  const end = station_(to);
  if (start === end) return [{ key: 'same-station', amount: 0, route: '同一駅' }];
  const options = new Map();
  for (const [origin, arrival, method, direction, raw, route, first, last] of rows) {
    const validFrom = dateValue_(first);
    const validThrough = dateValue_(last);
    const amount = String(raw ?? '').replace(/[,，]/g, '').trim();
    if (station_(origin) !== start || station_(arrival) !== end || method !== trip.method || direction !== trip.direction ||
        !validFrom || validFrom > trip.date || (last && (!validThrough || validThrough < trip.date)) || !/^[0-9]{1,7}$/.test(amount)) continue;
    const label = String(route ?? '').trim().slice(0, 180) || '登録経路';
    const key = Utilities.base64EncodeWebSafe(JSON.stringify([start, end, method, direction, amount, label, validFrom, validThrough]));
    options.set(key, { key, amount: Number(amount), route: label });
  }
  return [...options.values()].sort((a, b) => a.amount - b.amount || a.route.localeCompare(b.route, 'ja')).slice(0, 20);
}

function readChosenFare_(rows, from, to, trip, key, amount) {
  if (!key) return { amount: yen_(amount), route: '', source: '手入力' };
  const choice = fareOptions_(rows, from, to, trip).find((option) => option.key === key);
  if (!choice) throw new Error('運賃マスターが変更されました。運賃候補を取り直してください。');
  return { ...choice, source: choice.key === 'same-station' ? '同一駅' : '運賃マスター' };
}

function expenseStorage_() {
  const props = PropertiesService.getScriptProperties();
  const book = props.getProperty('EXPENSE_BOOK_ID');
  const folder = props.getProperty('EXPENSE_FOLDER_ID');
  if (!book || !folder) throw new Error('申請先が未設定です。管理者へご連絡ください。');
  return { book, folder };
}

function textValue_(value, max, label) {
  const result = typeof value === 'string' ? value.trim() : '';
  if (!result || result.length > max) throw new Error(label + 'を確認してください。');
  return result;
}

function dateValue_(value) {
  const match = String(value ?? '').trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!match) return '';
  const date = match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0');
  const timestamp = Date.parse(date + 'T00:00:00Z');
  return Number.isFinite(timestamp) && new Date(timestamp).toISOString().slice(0, 10) === date ? date : '';
}

function yen_(amount) {
  if (!/^\d{1,7}$/.test(String(amount ?? ''))) throw new Error('金額は0〜9,999,999円の整数で入力してください。');
  return Number(amount);
}

function expensePhoto_(encoded, filename) {
  if (typeof encoded !== 'string' || encoded.length > 4 * Math.ceil(EXPENSE.maxImage / 3) ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(encoded)) throw new Error('写真データを確認してください。');
  const bytes = Utilities.base64Decode(encoded);
  if (bytes.length < 4 || bytes.length > EXPENSE.maxImage ||
      (bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 || (bytes[2] & 255) !== 255 ||
      (bytes[bytes.length - 2] & 255) !== 255 || (bytes[bytes.length - 1] & 255) !== 217) {
    throw new Error('写真データを確認してください。');
  }
  return Utilities.newBlob(bytes, 'image/jpeg', filename);
}
