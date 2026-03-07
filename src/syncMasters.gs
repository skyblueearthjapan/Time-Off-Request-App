// ====== マスタ自動転記（元スプレッドシート → DB） ======

const SYNC_MAP = [
  { src: null, dst: SHEET.WORKER, srcKey: 'MASTER_SOURCE_WORKER_SHEET' },
];

/**
 * 作業員マスタを元スプレッドシート → DB へ同期（全置換）。
 * Settings: MASTER_SOURCE_SSID = 同期元スプレッドシートID（作業日報_全従業員用）
 *           MASTER_SOURCE_WORKER_SHEET = 同期元の作業員シート名
 * ※ 部署一覧は作業員マスタの「部署」列からユニーク取得するためM_DEPTは不要
 */
function syncAllMasters() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var settings = getSettings_();
    var sourceId = normalize_(settings['MASTER_SOURCE_SSID']);
    // フォールバック: 設定が空なら直接IDを使用
    if (!sourceId) {
      sourceId = '1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ';
      Logger.log('MASTER_SOURCE_SSID未設定のためデフォルトIDを使用: ' + sourceId);
    }

    var srcSS = SpreadsheetApp.openById(sourceId);
    var dstSS = getDb_();
    var log = [];

    for (var i = 0; i < SYNC_MAP.length; i++) {
      var m = SYNC_MAP[i];
      try {
        // 同期元シート名をSettingsから取得
        var srcSheetName = normalize_(settings[m.srcKey]);
        // フォールバック: 設定が空ならデフォルト値を使用
        if (!srcSheetName) {
          if (m.srcKey === 'MASTER_SOURCE_WORKER_SHEET') {
            srcSheetName = '作業員マスタ';
            Logger.log(m.srcKey + '未設定のためデフォルト値を使用: ' + srcSheetName);
          } else {
            log.push('SKIP ' + m.dst + ': ' + m.srcKey + ' が未設定');
            continue;
          }
        }

        var srcSh = srcSS.getSheetByName(srcSheetName);
        if (!srcSh) {
          log.push('SKIP ' + srcSheetName + ': 元シートなし');
          continue;
        }

        var dstSh = dstSS.getSheetByName(m.dst);
        if (!dstSh) {
          log.push('SKIP ' + m.dst + ': 転記先シートなし');
          continue;
        }

        var data = readSheetData_(srcSh);
        if (!data || !data.length) {
          log.push('SKIP ' + srcSheetName + ': 空');
          continue;
        }

        // 空行除外（1列目が空の行を除去、ヘッダは維持）
        data = compactRows_(data);

        // 通常の置換を試行、失敗時はシート再作成
        try {
          replaceMasterData_(dstSh, data);
        } catch (writeErr) {
          console.warn('replaceMasterData_ failed for ' + m.dst + ', recreating: ' + writeErr.message);
          dstSS.deleteSheet(dstSh);
          dstSh = dstSS.insertSheet(m.dst);
          dstSh.getRange(1, 1, data.length, data[0].length).setValues(data);
        }

        SpreadsheetApp.flush();
        log.push('OK ' + srcSheetName + ' -> ' + m.dst + ' (' + data.length + ' rows)');
      } catch (e) {
        log.push('ERROR ' + m.dst + ': ' + e.message);
      }
    }

    // カレンダーマスタも同期（別ソースSS）
    try {
      syncCalendarMaster();
      log.push('OK 社内カレンダーマスタ -> M_CALENDAR');
    } catch (calErr) {
      log.push('ERROR M_CALENDAR: ' + calErr.message);
    }

    var summary = 'syncAllMasters:\n' + log.join('\n');
    Logger.log(summary);
    console.log(summary);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 社内カレンダーマスタを外部スプレッドシート → M_CALENDAR へ同期（全置換）。
 * Settings: CALENDAR_SOURCE_SSID = 同期元スプレッドシートID（残業・休日出勤申請app）
 *           CALENDAR_SOURCE_SHEET = 同期元の社内カレンダーシート名
 */
function syncCalendarMaster() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var settings = getSettings_();
    var sourceId = normalize_(settings['CALENDAR_SOURCE_SSID']);
    if (!sourceId) {
      sourceId = '1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ';
      Logger.log('CALENDAR_SOURCE_SSID未設定のためデフォルトIDを使用: ' + sourceId);
    }
    var srcSheetName = normalize_(settings['CALENDAR_SOURCE_SHEET']);
    if (!srcSheetName) {
      srcSheetName = '社内カレンダーマスタ';
      Logger.log('CALENDAR_SOURCE_SHEET未設定のためデフォルト値を使用: ' + srcSheetName);
    }

    var srcSS = SpreadsheetApp.openById(sourceId);
    var srcSh = srcSS.getSheetByName(srcSheetName);
    if (!srcSh) throw new Error('元シートなし: ' + srcSheetName);

    var dstSh = requireSheet_(SHEET.CALENDAR);
    var data = readSheetData_(srcSh);
    if (!data || !data.length) throw new Error('社内カレンダーマスタが空です');

    data = compactRows_(data);

    try {
      replaceMasterData_(dstSh, data);
    } catch (writeErr) {
      console.warn('replaceMasterData_ failed for M_CALENDAR, recreating: ' + writeErr.message);
      var dstSS = getDb_();
      dstSS.deleteSheet(dstSh);
      dstSh = dstSS.insertSheet(SHEET.CALENDAR);
      dstSh.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    SpreadsheetApp.flush();
    var msg = 'syncCalendarMaster: OK (' + data.length + ' rows)';
    Logger.log(msg);
    console.log(msg);
  } finally {
    lock.releaseLock();
  }
}

// ====== ヘッダ行の実データ末尾から最終列を検出 ======
function detectLastColByHeader_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) return 0;
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var c = header.length; c >= 1; c--) {
    if (String(header[c - 1] || '').trim() !== '') return c;
  }
  return 1;
}

// ====== シートデータ読み取り（リトライ＋バッチフォールバック） ======
function readSheetData_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = detectLastColByHeader_(sh);
  if (lastRow === 0 || lastCol === 0) return null;

  // 1st try
  try {
    return sh.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (e) {
    console.warn('readSheetData_ 1st try failed: ' + e.message);
  }

  // 2nd try with delay
  Utilities.sleep(2000);
  try {
    return sh.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (e) {
    console.warn('readSheetData_ 2nd try failed: ' + e.message);
  }

  // 3rd: batch read (50 rows at a time)
  Utilities.sleep(2000);
  try {
    var BATCH = 50;
    var all = [];
    for (var r = 1; r <= lastRow; r += BATCH) {
      var n = Math.min(BATCH, lastRow - r + 1);
      var batch = sh.getRange(r, 1, n, lastCol).getValues();
      for (var i = 0; i < batch.length; i++) all.push(batch[i]);
    }
    return all;
  } catch (e) {
    throw new Error('Read failed after all retries: ' + e.message);
  }
}

// ====== シート全置換（値のみ） ======
function replaceMasterData_(sh, values) {
  var rows = values.length;
  var cols = values[0].length;

  sh.clearContents();

  if (sh.getMaxRows() < rows) {
    sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  }
  if (sh.getMaxColumns() < cols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  }

  var excessRows = sh.getMaxRows() - rows;
  if (excessRows > 5) {
    sh.deleteRows(rows + 1, excessRows);
  }
  var excessCols = sh.getMaxColumns() - cols;
  if (excessCols > 2) {
    sh.deleteColumns(cols + 1, excessCols);
  }

  sh.getRange(1, 1, rows, cols).setValues(values);
}

// ====== 空行除外 ======
function compactRows_(values) {
  var header = values[0];
  var body = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() !== '') {
      body.push(values[i]);
    }
  }
  return [header].concat(body);
}
