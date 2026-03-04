// ====== マスタ自動転記（元スプレッドシート → DB） ======

const SYNC_MAP = [
  { src: null, dst: SHEET.DEPT, srcKey: 'MASTER_SOURCE_DEPT_SHEET' },
  { src: null, dst: SHEET.WORKER, srcKey: 'MASTER_SOURCE_WORKER_SHEET' },
];

/**
 * 全マスタを元スプレッドシート → DB へ同期（全置換）。
 * Settings: MASTER_SOURCE_SSID = 同期元スプレッドシートID
 *           MASTER_SOURCE_DEPT_SHEET = 部署シート名
 *           MASTER_SOURCE_WORKER_SHEET = 作業員シート名
 */
function syncAllMasters() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var settings = getSettings_();
    var sourceId = normalize_(settings['MASTER_SOURCE_SSID']);
    if (!sourceId) {
      throw new Error(
        'Settings に MASTER_SOURCE_SSID が未設定です。\n' +
        '同期元スプレッドシートのIDを M_SYSTEM_SETTING シートに追加してください。'
      );
    }

    var srcSS = SpreadsheetApp.openById(sourceId);
    var dstSS = getDb_();
    var log = [];

    for (var i = 0; i < SYNC_MAP.length; i++) {
      var m = SYNC_MAP[i];
      try {
        // 同期元シート名をSettingsから取得
        var srcSheetName = normalize_(settings[m.srcKey]);
        if (!srcSheetName) {
          log.push('SKIP ' + m.dst + ': ' + m.srcKey + ' が未設定');
          continue;
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

    var summary = 'syncAllMasters:\n' + log.join('\n');
    Logger.log(summary);
    console.log(summary);
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
