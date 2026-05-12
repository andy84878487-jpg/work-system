/**
 * 統晉企業 · 系統 ↔ Google Sheets/Drive 同步 v2
 *
 * 變更內容（相對於 v1）：
 *   - handleDeletes 回傳新增 deletedIds / notFoundIds 兩個陣列，
 *     APP 端依此精確判定每一筆 dirty 是否要清。
 *   - 對應 APP 端 pushToSheet 的新版判斷邏輯（v2 之後的測試系統 _22）。
 *
 * 同步原則：
 *   - upsert 模式：主鍵存在 → 更新；不存在 → 追加
 *   - delete 指令：APP 刪除資料時會送 deletes 清單，Apps Script 從試算表刪除對應列
 *     回傳的 deletedIds：實際刪到的 id；notFoundIds：找不到對應列的 id（已處理但無事可做）
 *   - 「職員名單」單向：APP 只能「拉取」，不能推送（保護個資）
 *   - 試算表上手動填的資料會被保留（除非主鍵碰撞或被 delete 指令命中）
 *   - 電話/聯絡資訊/機號 等欄位會被強制為「文字」格式，避免前導 0 被吃掉
 */

// ─── CONFIG ──────────────────────────────────────────────────────
const DRIVE_FOLDER_ID = '1paZK2FMV1TwCMewKXmpD2i8U1TzqU2Tt';

// 主鍵欄位設定
const PRIMARY_KEYS = {
  '客戶公司': '公司ID',
  '客戶資料': '客戶名稱',
  '聯絡人': '聯絡人ID',
  '機器清單': '機號',
  '合約': '合約ID',
  '簽單': null,                // 複合主鍵：單號+序號（刪除時用單號）
  '製冰機 報價': null,         // 複合主鍵：廠牌+機型
  '材料 報價': '名稱',
  '維修 報價': '名稱',
  '收款': '單號',
};

// 黑名單：不接受推送（保護個資）
const PUSH_BLACKLIST = ['職員名單'];

// 強制 TEXT 格式的欄位名稱關鍵字（避免「0050」被當數字而變成「50」）
const TEXT_FORMAT_KEYWORDS = [
  // 電話/聯絡相關
  '電話','聯絡資訊',
  // 識別碼類 — 防止前導 0 被吃掉
  '機號','序號','PIN','統一編號','身分證',
  '匯款帳號','合約ID','單號','客戶ID','公司ID','聯絡人ID',
  '收款ID','發票號碼',
  // 名稱類 — 客戶/公司/品名可能會用純數字命名
  '名稱','姓名','客戶名稱','公司名稱','店家名稱','客戶','公司','店家',
  // 地址（門牌可能含 0 開頭）
  '地址',
  // 內容類
  '品名','備註','服務類型','類型','狀態','付款狀態','收款方式',
];

function _isTextColumn(headerName) {
  const h = String(headerName||'');
  return TEXT_FORMAT_KEYWORDS.some(k => h.indexOf(k) >= 0);
}

// 設定指定列的特定欄位為文字格式 (個別 try/catch，避免某欄有資料驗證時整個 push 失敗)
function _ensureTextFormat(sheet, headers, rowNum) {
  for (let c = 0; c < headers.length; c++) {
    if (_isTextColumn(headers[c])) {
      try { sheet.getRange(rowNum, c+1).setNumberFormat('@'); } catch (e) {}
    }
  }
}

// ─── doGet ───────────────────────────────────────────────────────
function doGet(e) {
  return _json({
    ok: true,
    message: '統晉企業同步服務運作中',
    timestamp: new Date().toISOString()
  });
}

// ─── doPost ──────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const mode = body.mode || 'upsert';

    if (mode === 'ping') {
      return _json({ ok:true, message:'pong' });
    }

    if (mode === 'pull') {
      return _json(handlePull(body.table));
    }

    // 拉取所有分頁的資料（給 APP 端做雙向同步用）
    if (mode === 'pullAll') {
      return _json(handlePullAll());
    }

    if (mode === 'fullSync' || mode === 'delta' || mode === 'upsert') {
      // ⚠️ Upsert 失敗不能擋掉 Delete - 兩段獨立 try
      let upsertResult, upsertErr;
      try {
        upsertResult = handleUpsert(body.data || {});
      } catch (e1) {
        upsertErr = String(e1);
        upsertResult = { ok:false, upsertError: upsertErr };
      }

      let deleteResult, deleteErr;
      if (body.deletes && body.deletes.length > 0) {
        try {
          deleteResult = handleDeletes(body.deletes);
        } catch (e2) {
          deleteErr = String(e2);
          deleteResult = { _error: deleteErr };
        }
      }

      // 回傳結構 - ok 設成 true，個別錯誤放在 errors 內，讓 APP 可以個別處理
      const finalResult = {
        ok: true,
        mode: upsertResult.mode || mode,
        summary: upsertResult.summary || {},
        timestamp: new Date().toISOString()
      };
      if (deleteResult) finalResult.deleteResult = deleteResult;
      if (upsertErr) finalResult.upsertError = upsertErr;
      if (deleteErr) finalResult.deleteError = deleteErr;
      return _json(finalResult);
    }

    return _json({ ok:false, error:'unknown mode: ' + mode });
  } catch (err) {
    return _json({ ok:false, error: String(err), stack: err.stack ? err.stack.slice(0, 500) : '' });
  }
}

// ─── 拉取試算表 → APP ───────────────────────────────────────────
function handlePull(table) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (table === '職員名單') {
    const sheet = ss.getSheetByName('職員名單');
    if (!sheet) return { ok:true, staff: [] };
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return { ok:true, staff: [] };
    const headers = rows[0].map(h => String(h).trim());
    const staff = [];
    for (let i = 1; i < rows.length; i++) {
      const r = {};
      headers.forEach((h, j) => r[h] = rows[i][j]);
      const name = String(r['姓名']||'').trim();
      if (!name) continue;
      // ⚠️ 只回傳姓名 + PIN + 角色（其他個資不出試算表）
      // 角色：管理員 / 帳務員 / 工程師（預設）
      const roleStr = String(r['角色']||'').trim();
      let role = 'staff';
      if (roleStr === '管理員') role = 'admin';
      else if (roleStr === '帳務員' || roleStr === '帳務') role = 'accounting';
      staff.push({
        name: name,
        pin: String(r['PIN']||'').trim(),
        role: role
      });
    }
    return { ok:true, staff };
  }

  return { ok:false, error:'unknown table: '+table };
}

// 拉取所有分頁（雙向同步用）
function handlePullAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  // 要拉的分頁清單（排除職員名單，那個用 handlePull 拉）
  const tables = ['客戶公司','客戶資料','聯絡人','機器清單','合約','簽單','製冰機 報價','材料 報價','維修 報價','收款'];

  tables.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { result[name] = []; return; }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { result[name] = []; return; }
    const lastCol = sheet.getLastColumn();
    const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = all[0].map(h => String(h).trim());
    const rows = [];
    for (let i = 1; i < all.length; i++) {
      const r = {};
      let hasContent = false;
      headers.forEach((h, j) => {
        const v = all[i][j];
        // 把 Date 物件轉字串
        let val = v;
        if (v instanceof Date) {
          val = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        }
        r[h] = val;
        if (val !== '' && val !== null && val !== undefined) hasContent = true;
      });
      if (hasContent) rows.push(r);
    }
    result[name] = rows;
  });

  // 順便也拉職員名單（同 handlePull 邏輯但全部資料）
  const staffSheet = ss.getSheetByName('職員名單');
  if (staffSheet) {
    const sRows = staffSheet.getDataRange().getValues();
    if (sRows.length >= 2) {
      const sHeaders = sRows[0].map(h => String(h).trim());
      const staff = [];
      for (let i = 1; i < sRows.length; i++) {
        const r = {};
        sHeaders.forEach((h, j) => r[h] = sRows[i][j]);
        const name = String(r['姓名']||'').trim();
        if (!name) continue;
        const roleStr = String(r['角色']||'').trim();
        let role = 'staff';
        if (roleStr === '管理員') role = 'admin';
        else if (roleStr === '帳務員' || roleStr === '帳務') role = 'accounting';
        staff.push({
          name: name,
          pin: String(r['PIN']||'').trim(),
          role: role
        });
      }
      result['職員名單'] = staff;
    }
  }

  return { ok:true, data: result, timestamp: new Date().toISOString() };
}

// ─── Upsert 同步 ──────────────────────────────────────────
function handleUpsert(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = {};

  for (const sheetName in data) {
    try {
      if (PUSH_BLACKLIST.indexOf(sheetName) >= 0) {
        summary[sheetName] = { skipped: 'blacklisted' };
        continue;
      }

      const rows = data[sheetName] || [];
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        summary[sheetName] = { skipped: 'sheet not found' };
        continue;
      }

      if (rows.length === 0) {
        summary[sheetName] = { added: 0, updated: 0 };
        continue;
      }

      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

      // 預先把整欄文字欄位設為 @ 格式（一次性，後續 append 也會繼承）
      for (let c = 0; c < headers.length; c++) {
        if (_isTextColumn(headers[c])) {
          try { sheet.getRange(1, c+1, sheet.getMaxRows(), 1).setNumberFormat('@'); } catch (e) {}
        }
      }

      if (sheetName === '簽單') {
        _upsertCompositeKey(sheet, headers, rows, ['單號','序號'], summary, sheetName);
        continue;
      }
      if (sheetName === '製冰機 報價') {
        _upsertCompositeKey(sheet, headers, rows, ['廠牌','機型'], summary, sheetName);
        continue;
      }

      const pkField = PRIMARY_KEYS[sheetName];
      if (!pkField) {
        _appendOnly(sheet, headers, rows, summary, sheetName);
        continue;
      }

      _upsertSingleKey(sheet, headers, rows, pkField, summary, sheetName);
    } catch (sheetErr) {
      // 單張 sheet 出錯不影響其他分頁
      summary[sheetName] = { error: String(sheetErr).slice(0, 200) };
    }
  }

  return { ok:true, mode:'upsert', summary, timestamp:new Date().toISOString() };
}

// ─── Delete 處理（v2：回傳 deletedIds / notFoundIds）────────
function handleDeletes(deletes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  // 為每個 table 初始化 result 結構
  const ensureBucket = (name) => {
    if (!result[name]) {
      result[name] = { deleted: 0, deletedIds: [], notFoundIds: [] };
    }
    return result[name];
  };

  for (const del of deletes) {
    const sheetName = del.table;
    const id = String(del.id||'').trim();
    if (!sheetName || !id) continue;

    const bucket = ensureBucket(sheetName);

    if (PUSH_BLACKLIST.indexOf(sheetName) >= 0) {
      bucket.blacklisted = (bucket.blacklisted||0) + 1;
      continue;
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      bucket.error = 'sheet not found';
      continue;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      // 表是空的 → 視為找不到（已處理）
      bucket.notFoundIds.push(id);
      continue;
    }

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    let thisDeleted = 0;

    // ⚠️ 用 getDisplayValues() 取出「顯示字串」(避免 Date/Number 字串化跟本地 id 對不上)
    //    再搭配 getValues() 雙重比對，提升找列的成功率
    const displayValues = sheet.getRange(2, 1, lastRow-1, lastCol).getDisplayValues();
    const rawValues     = sheet.getRange(2, 1, lastRow-1, lastCol).getValues();

    // 兩種比對：trim 後字串相等、或全部去除非英數字後相等 (容忍空白/破折號)
    const _norm = s => String(s||'').trim();
    const _looseNorm = s => String(s||'').replace(/[^0-9A-Za-z一-鿿]/g,'').toLowerCase();
    const matchCell = (rawCell, dispCell, target) => {
      const tn = _norm(target);
      if (!tn) return false;
      const r = _norm(rawCell);
      const d = _norm(dispCell);
      if (r === tn || d === tn) return true;
      const tlh = _looseNorm(target);
      if (tlh && (_looseNorm(r) === tlh || _looseNorm(d) === tlh)) return true;
      return false;
    };

    // 安全刪除：如果是最後一列無法 deleteRow（Sheets 不允許刪除所有非凍結列）→ 改清空欄位
    const safeDeleteRow = (rowNum, headersLen) => {
      try {
        sheet.deleteRow(rowNum);
        return true;
      } catch (e1) {
        // 退路：把該列所有欄位清空（視為已刪除）
        try {
          const blank = new Array(headersLen).fill('');
          sheet.getRange(rowNum, 1, 1, headersLen).setValues([blank]);
          return true;
        } catch (e2) {
          return false;
        }
      }
    };

    if (sheetName === '簽單') {
      const noColIdx = headers.indexOf('單號');
      if (noColIdx < 0) { bucket.error = 'no 單號 column'; continue; }
      for (let i = displayValues.length - 1; i >= 0; i--) {
        if (matchCell(rawValues[i][noColIdx], displayValues[i][noColIdx], id)) {
          if (safeDeleteRow(i + 2, headers.length)) thisDeleted++;
        }
      }
    }
    else if (sheetName === '製冰機 報價') {
      const parts = id.split('|');
      if (parts.length !== 2) {
        bucket.notFoundIds.push(id);
        continue;
      }
      const brand = parts[0], model = parts[1];
      const bIdx = headers.indexOf('廠牌');
      const mIdx = headers.indexOf('機型');
      if (bIdx < 0 || mIdx < 0) {
        bucket.error = 'no 廠牌/機型 column';
        continue;
      }
      for (let i = displayValues.length - 1; i >= 0; i--) {
        if (matchCell(rawValues[i][bIdx], displayValues[i][bIdx], brand) &&
            matchCell(rawValues[i][mIdx], displayValues[i][mIdx], model)) {
          if (safeDeleteRow(i + 2, headers.length)) thisDeleted++;
        }
      }
    }
    else {
      const pkField = PRIMARY_KEYS[sheetName];
      if (!pkField) {
        bucket.error = 'no PK config';
        continue;
      }
      const pkColIdx = headers.indexOf(pkField);
      if (pkColIdx < 0) {
        bucket.error = 'PK column not found: ' + pkField;
        continue;
      }
      for (let i = displayValues.length - 1; i >= 0; i--) {
        if (matchCell(rawValues[i][pkColIdx], displayValues[i][pkColIdx], id)) {
          if (safeDeleteRow(i + 2, headers.length)) thisDeleted++;
        }
      }
    }

    bucket.deleted += thisDeleted;
    if (thisDeleted > 0) bucket.deletedIds.push(id);
    else bucket.notFoundIds.push(id);
  }

  return result;
}

// ─── Upsert helpers ─────────────────────────────────────────────
function _upsertSingleKey(sheet, headers, rows, pkField, summary, sheetName) {
  const lastRow = sheet.getLastRow();
  const pkColIdx = headers.indexOf(pkField);
  if (pkColIdx < 0) {
    summary[sheetName] = { error: 'PK column not found: ' + pkField };
    return;
  }

  // 第一步：掃描現有列，建立 PK → 列號的索引；若同 PK 多筆，從下往上刪除只留首筆
  const existingPks = {};
  let dedupRemoved = 0;
  if (lastRow >= 2) {
    const pkRange = sheet.getRange(2, pkColIdx+1, lastRow-1, 1).getValues();
    const dupRows = [];
    pkRange.forEach((row, i) => {
      const k = String(row[0]||'').trim();
      if (!k) return;
      if (existingPks[k]) {
        dupRows.push(i + 2);
      } else {
        existingPks[k] = i + 2;
      }
    });
    for (let i = dupRows.length - 1; i >= 0; i--) {
      sheet.deleteRow(dupRows[i]);
      dedupRemoved++;
    }
    if (dedupRemoved > 0) {
      const newLast = sheet.getLastRow();
      Object.keys(existingPks).forEach(k => delete existingPks[k]);
      if (newLast >= 2) {
        const newRange = sheet.getRange(2, pkColIdx+1, newLast-1, 1).getValues();
        newRange.forEach((row, i) => {
          const k = String(row[0]||'').trim();
          if (k && !existingPks[k]) existingPks[k] = i + 2;
        });
      }
    }
  }

  let added = 0, updated = 0, skipped = 0, errors = 0;
  rows.forEach(rec => {
    const k = String(rec[pkField]||'').trim();
    if (!k) { skipped++; return; }

    const rowValues = headers.map(h => _formatVal(rec[h], h));
    try {
      if (existingPks[k]) {
        sheet.getRange(existingPks[k], 1, 1, headers.length).setValues([rowValues]);
        _ensureTextFormat(sheet, headers, existingPks[k]);
        updated++;
      } else {
        sheet.appendRow(rowValues);
        _ensureTextFormat(sheet, headers, sheet.getLastRow());
        existingPks[k] = sheet.getLastRow();
        added++;
      }
    } catch (e) {
      errors++;
    }
  });

  summary[sheetName] = { added: added, updated: updated, skipped: skipped, errors: errors, dedupRemoved: dedupRemoved };
}

function _upsertCompositeKey(sheet, headers, rows, pkFields, summary, sheetName) {
  const lastRow = sheet.getLastRow();
  const pkColIdxs = pkFields.map(f => headers.indexOf(f));
  if (pkColIdxs.some(i => i < 0)) {
    summary[sheetName] = { error: 'PK column not found' };
    return;
  }

  const existingPks = {};
  let dedupRemoved = 0;
  if (lastRow >= 2) {
    const allRange = sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn()).getValues();
    const dupRows = [];
    allRange.forEach((row, i) => {
      const k = pkColIdxs.map(idx => String(row[idx]||'').trim()).join('|');
      if (!k || k.match(/^\|+$/)) return;
      if (existingPks[k]) {
        dupRows.push(i + 2);
      } else {
        existingPks[k] = i + 2;
      }
    });
    for (let i = dupRows.length - 1; i >= 0; i--) {
      sheet.deleteRow(dupRows[i]);
      dedupRemoved++;
    }
    if (dedupRemoved > 0) {
      const newLast = sheet.getLastRow();
      Object.keys(existingPks).forEach(k => delete existingPks[k]);
      if (newLast >= 2) {
        const newRange = sheet.getRange(2, 1, newLast-1, sheet.getLastColumn()).getValues();
        newRange.forEach((row, i) => {
          const k = pkColIdxs.map(idx => String(row[idx]||'').trim()).join('|');
          if (k && !k.match(/^\|+$/) && !existingPks[k]) existingPks[k] = i + 2;
        });
      }
    }
  }

  let added = 0, updated = 0, skipped = 0, errors = 0;
  rows.forEach(rec => {
    const k = pkFields.map(f => String(rec[f]||'').trim()).join('|');
    if (!k || k.match(/^\|+$/)) { skipped++; return; }

    const rowValues = headers.map(h => _formatVal(rec[h], h));
    try {
      if (existingPks[k]) {
        sheet.getRange(existingPks[k], 1, 1, headers.length).setValues([rowValues]);
        _ensureTextFormat(sheet, headers, existingPks[k]);
        updated++;
      } else {
        sheet.appendRow(rowValues);
        _ensureTextFormat(sheet, headers, sheet.getLastRow());
        existingPks[k] = sheet.getLastRow();
        added++;
      }
    } catch (e) {
      errors++;
    }
  });

  summary[sheetName] = { added: added, updated: updated, skipped: skipped, errors: errors, dedupRemoved: dedupRemoved };
}

function _appendOnly(sheet, headers, rows, summary, sheetName) {
  let errors = 0;
  rows.forEach(rec => {
    const rowValues = headers.map(h => _formatVal(rec[h], h));
    try {
      sheet.appendRow(rowValues);
      _ensureTextFormat(sheet, headers, sheet.getLastRow());
    } catch (e) { errors++; }
  });
  summary[sheetName] = { appended: rows.length - errors, errors: errors };
}

// _formatVal：把值轉成試算表能寫的格式
// 對於文字類欄位（電話/機號等），純數字字串會加上「'」前綴
// 這是 Google Sheets 的特殊用法：'0912 會被儲存為文字 "0912"（顯示時不顯示 '）
function _formatVal(v, header) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (header && _isTextColumn(header)) {
    const s = String(v);
    if (!s) return '';
    // 純數字（含逗號分隔的多筆電話、空格、+ - .）→ 加 ' 強制當文字
    if (/^[\d\s\+\-\.,()]+$/.test(s)) {
      return "'" + s;
    }
    return s;
  }
  return v;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── 測試 ─────────────────────────────────────────────────────────
function testPing() {
  const r = doGet({});
  Logger.log(r.getContent());
}

// 測試 delete 流程（手動執行用）
function testDelete() {
  const result = handleDeletes([
    { table:'簽單', id:'W20260513001', op:'delete' }
  ]);
  Logger.log(JSON.stringify(result, null, 2));
}
