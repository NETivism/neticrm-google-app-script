/**
 * netiCRM 捐款日期區間同步至 Google Sheet
 *
 * 使用方式：
 *  1. 執行一次 setupProperties() 填入你的站台網址、API 金鑰與自訂欄位 ID
 *  2. 執行一次 createDailyTrigger() 安裝每日自動排程（可指定執行時間）
 *  3. 若需要手動測試，執行 syncDatesManual() 輸入起訖日，或直接呼叫
 *     syncDates('2026-02-01', '2026-02-28')
 *
 * 欄位順序（netiCRM捐款同步表格）：
 *  A 金流機制 | B 費用類型 | C 捐款來源 | D 定期捐款編號 | E 總金額
 *  F 收到日期 | G 收據開立日期 | H 捐款編號（唯一鍵）| I 交易編號
 *  J 收據編號 | K 捐款狀態 | L 付款工具
 *  M 姓氏 | N 名字 | O 收據抬頭／姓名 | P 捐款徵信名稱 | Q 報稅憑證
 */

// ─── 自訂選單 ─────────────────────────────────────────────────────────────────

/** 試算表開啟時自動建立「netiCRM 同步」選單。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('netiCRM 同步')
    .addItem('① 安裝設定', 'setupProperties')
    .addItem('② 設定每日自動排程時間', 'createDailyTrigger')
    .addItem('③ 同步指定日期區間', 'syncDatesManual')
    .addToUi();
}

// ─── 安全憑證設定 ────────────────────────────────────────────────────────────

/**
 * 初始化腳本屬性（執行一次即可）。
 * 完成後可至「專案設定 → 指令碼屬性」確認或修改。
 * 注意：此函式執行後金鑰即儲存在 GCP，原始碼中不含明文金鑰。
 *
 * 自訂欄位 ID 查詢方式：
 *  netiCRM 後台 → 管理 → 自訂欄位 → 找到對應欄位後查看 URL 中的 id 參數
 *  或呼叫：POST <entrypoint>?entity=CustomField&action=get，body 為 {}
 */
function setupProperties() {
  const props = PropertiesService.getScriptProperties();
  const ui    = SpreadsheetApp.getUi();

  // 欄位定義：key、提示說明、驗證類型、是否必填
  const fields = [
    {
      key:      'NETICRM_BASE_URL',
      label:    '站台 API 網址',
      hint:     '必須以 https:// 開頭，例如：https://example.org/..../civicrm/extern/rest.php',
      type:     'url',
      required: true,
    },
    {
      key:      'NETICRM_API_KEY',
      label:    '聯絡人 API Key',
      hint:     '只允許英文字母與數字',
      type:     'alphanumeric',
      required: true,
    },
    {
      key:      'NETICRM_SITE_KEY',
      label:    '站台金鑰（Site Key）',
      hint:     '只允許英文字母與數字',
      type:     'alphanumeric',
      required: true,
    },
    {
      key:      'CUSTOM_FIELD_RECEIPT',
      label:    '收據抬頭／姓名 自訂欄位 ID',
      hint:     '只允許數字，不需要請留空',
      type:     'number',
      required: false,
    },
    {
      key:      'CUSTOM_FIELD_CREDIT',
      label:    '捐款徵信名稱 自訂欄位 ID',
      hint:     '只允許數字，不需要請留空',
      type:     'number',
      required: false,
    },
    {
      key:      'CUSTOM_FIELD_TAX_CERT',
      label:    '報稅憑證 自訂欄位 ID',
      hint:     '只允許數字，不需要請留空',
      type:     'number',
      required: false,
    },
  ];

  let changed = 0;

  for (let i = 0; i < fields.length; i++) {
    const field = fields[i];

    if (props.getProperty(field.key) !== null) {
      Logger.log('Skipping existing property: ' + field.key);
      continue;
    }

    // 反覆提示，直到輸入合法或使用者取消
    let saved = false;
    while (!saved) {
      const resp = ui.prompt(
        '設定屬性 (' + (i + 1) + '/' + fields.length + ')：' + field.label,
        field.hint,
        ui.ButtonSet.OK_CANCEL
      );

      if (resp.getSelectedButton() !== ui.Button.OK) {
        Logger.log('Skipped by user: ' + field.key);
        break;
      }

      const input = resp.getResponseText().trim();

      // 空值處理
      if (input === '') {
        if (field.required) {
          ui.alert('此欄位為必填，請輸入值。');
          continue;
        }
        props.setProperty(field.key, '');
        changed++;
        Logger.log('Set property (empty): ' + field.key);
        saved = true;
        continue;
      }

      // 格式驗證
      let errorMsg = '';
      if (field.type === 'url' && !/^https:\/\/.+/.test(input)) {
        errorMsg = '網址必須以 https:// 開頭。';
      } else if (field.type === 'alphanumeric' && !/^[A-Za-z0-9]+$/.test(input)) {
        errorMsg = '只允許英文字母與數字，不可包含空格或特殊符號。';
      } else if (field.type === 'number' && !/^\d+$/.test(input)) {
        errorMsg = '只允許數字（正整數）。';
      }

      if (errorMsg) {
        ui.alert('格式錯誤：' + errorMsg);
        continue;
      }

      props.setProperty(field.key, input);
      changed++;
      Logger.log('Set property: ' + field.key + ' = ' + input);
      saved = true;
    }
  }

  if (changed > 0) {
    Logger.log('Script properties saved (' + changed + ' new).');
    ui.alert('設定完成！共儲存 ' + changed + ' 項屬性。\n\n如需日後修改，請至「專案設定 → 指令碼屬性」。');
  } else {
    Logger.log('All properties already set, nothing changed.');
    ui.alert('所有屬性皆已設定，無需變更。\n\n如需修改，請至「專案設定 → 指令碼屬性」。');
  }
}

/** 從 PropertiesService 讀取設定（不含明文金鑰）。 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();

  function customKey(id) {
    return id ? 'custom_' + id : null;
  }

  return {
    baseUrl:   props.getProperty('NETICRM_BASE_URL'),
    apiKey:    props.getProperty('NETICRM_API_KEY'),
    siteKey:   props.getProperty('NETICRM_SITE_KEY'),
    sheetName: 'netiCRM捐款同步表格',
    pageSize:  100,  // API 單次最大筆數為 100
    fieldMap: {
      receipt_name: customKey(props.getProperty('CUSTOM_FIELD_RECEIPT')),
      credit_name:  customKey(props.getProperty('CUSTOM_FIELD_CREDIT')),
      tax_cert:     customKey(props.getProperty('CUSTOM_FIELD_TAX_CERT')),
    },
  };
}

// ─── 觸發器 ───────────────────────────────────────────────────────────────────

/**
 * 設定每日自動排程時間（Asia/Taipei），每次執行會自動抓取前一日捐款。
 * 若已存在排程，會先提示目前設定的時間，確認後才允許重新設定。
 */
function createDailyTrigger() {
  const props = PropertiesService.getScriptProperties();
  const ui    = SpreadsheetApp.getUi();

  // 檢查是否已有排程
  const existing = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyRun');
  const savedHour = props.getProperty('DAILY_TRIGGER_HOUR');

  if (existing.length > 0) {
    const currentTime = savedHour !== null ? savedHour + ':00' : '（時間未知）';
    const confirm = ui.alert(
      '排程已存在',
      '目前每日排程設定於 ' + currentTime + ' (Asia/Taipei)。\n\n是否要重新設定執行時間？',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  // 提示輸入小時（0–23）
  const resp = ui.prompt(
    '設定每日排程時間',
    '請輸入每日自動執行的小時（0–23，24 小時制，Asia/Taipei）：',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const input = resp.getResponseText().trim();
  if (!/^\d{1,2}$/.test(input) || parseInt(input, 10) < 0 || parseInt(input, 10) > 23) {
    ui.alert('格式錯誤：請輸入 0 到 23 之間的整數。');
    return;
  }

  const hour = parseInt(input, 10);

  // 移除舊排程
  existing.forEach(t => ScriptApp.deleteTrigger(t));

  // 建立新排程
  ScriptApp.newTrigger('dailyRun')
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .inTimezone('Asia/Taipei')
    .create();

  props.setProperty('DAILY_TRIGGER_HOUR', String(hour));
  Logger.log('Daily trigger created: dailyRun at ' + hour + ':00 Asia/Taipei');
  ui.alert('排程設定完成！每日 ' + hour + ':00 (Asia/Taipei) 將自動同步前一日捐款。');
}

// ─── 主流程 ───────────────────────────────────────────────────────────────────

/** 排程進入點：抓取昨日捐款。 */
function dailyRun() {
  const yesterday = Utilities.formatDate(
    new Date(Date.now() - 86400000),
    'Asia/Taipei',
    'yyyy-MM-dd'
  );
  syncDates(yesterday, yesterday);
}

/**
 * 手動除錯用：透過對話框輸入起訖日後執行同步。
 * 在 Apps Script 編輯器中執行此函式，會彈出輸入視窗。
 */
function syncDatesManual() {
  const ui = SpreadsheetApp.getUi();

  const startResp = ui.prompt(
    '手動同步捐款 (1/2)',
    '請輸入開始日期（格式：yyyy-MM-dd，例如 2026-02-01）：',
    ui.ButtonSet.OK_CANCEL
  );
  if (startResp.getSelectedButton() !== ui.Button.OK) return;

  const endResp = ui.prompt(
    '手動同步捐款 (2/2)',
    '請輸入結束日期（格式：yyyy-MM-dd，例如 2026-02-28）：',
    ui.ButtonSet.OK_CANCEL
  );
  if (endResp.getSelectedButton() !== ui.Button.OK) return;

  const startDate = startResp.getResponseText().trim();
  const endDate   = endResp.getResponseText().trim();
  const dateRe    = /^\d{4}-\d{2}-\d{2}$/;

  if (!dateRe.test(startDate) || !dateRe.test(endDate)) {
    ui.alert('日期格式錯誤，請使用 yyyy-MM-dd。');
    return;
  }
  if (startDate > endDate) {
    ui.alert('開始日期不可晚於結束日期。');
    return;
  }

  syncDates(startDate, endDate);
  ui.alert('同步完成！請查看執行紀錄（View → Logs）了解詳情。');
}

/**
 * 同步指定日期區間的全部捐款至 netiCRM捐款同步表格。
 * 結束日自動涵蓋至當天 23:59:59。
 *
 * @param {string} startDate  格式 'yyyy-MM-dd'，例如 '2026-02-01'
 * @param {string} endDate    格式 'yyyy-MM-dd'，例如 '2026-02-28'
 */
function syncDates(startDate, endDate) {
  const config = getConfig();
  if (!config.baseUrl || !config.apiKey || !config.siteKey) {
    throw new Error('請先執行 setupProperties() 填入站台設定。');
  }

  Logger.log('=== 開始同步 [' + startDate + ' ~ ' + endDate + '] ===');
  Logger.log('Custom field map: ' + JSON.stringify(config.fieldMap));

  // 1. 查詢筆數
  const total = getContributionCount(config, startDate, endDate);
  Logger.log('Total contributions: ' + total);
  if (total === 0) {
    Logger.log('No contributions found, skipping.');
    return;
  }

  // 2. 預取選項 label 對應表（contribution_type_id / contribution_status_id / payment_instrument_id）
  const optionLabels = fetchOptionLabels(config);

  // 3. 分頁抓取全部捐款
  const contributions = fetchAllContributions(config, startDate, endDate, total, optionLabels);
  Logger.log('Fetched ' + contributions.length + ' contributions.');

  // 4. 寫入 Sheet
  writeToSheet(config, contributions, startDate, endDate);
  Logger.log('=== 同步完成 [' + startDate + ' ~ ' + endDate + '] ===');
}

// ─── 取得捐款筆數 ─────────────────────────────────────────────────────────────

/**
 * 呼叫 Contribution.getcount 取得日期區間的捐款總筆數。
 * 回應格式：{ "is_error": 0, "result": N }
 *
 * @param {Object} config
 * @param {string} startDate  'yyyy-MM-dd'
 * @param {string} endDate    'yyyy-MM-dd'
 * @returns {number}
 */
function getContributionCount(config, startDate, endDate) {
  const result = callApi(config, 'Contribution', 'getcount', buildDateRange(startDate, endDate));
  return result.result || 0;
}

// ─── 分頁抓取捐款 ─────────────────────────────────────────────────────────────

/**
 * 欄位定義表（單一來源）。
 * 新增／移除欄位只需修改此處，return 清單、資料解析、表頭均自動同步。
 *
 * source:
 *   'contribution' — 直接從捐款物件讀取，apiField 為 API 欄位名稱
 *   'contact'      — 從 api.Contact.get 結果讀取，apiField 為聯絡人欄位名稱
 *   'custom'       — 自訂欄位，customKey 對應 config.fieldMap 的 key
 * format: true — 值需經 formatDate() 處理
 */
const FIELD_DEFS = [
  { header: '金流機制',      source: 'contribution', apiField: 'payment_processor_id' },
  { header: '費用類型',      source: 'contribution', apiField: 'contribution_type_id',   options: true },
  { header: '捐款來源',      source: 'contribution', apiField: 'contribution_source' },
  { header: '定期捐款編號',  source: 'contribution', apiField: 'contribution_recur_id' },
  { header: '總金額',        source: 'contribution', apiField: 'total_amount' },
  { header: '收到日期',      source: 'contribution', apiField: 'receive_date',        format: true },
  { header: '收據開立日期',  source: 'contribution', apiField: 'receipt_date',        format: true },
  { header: '捐款編號',      source: 'contribution', apiField: 'contribution_id' },
  { header: '交易編號',      source: 'contribution', apiField: 'trxn_id' },
  { header: '收據編號',      source: 'contribution', apiField: 'receipt_id' },
  { header: '捐款狀態',      source: 'contribution', apiField: 'contribution_status_id', options: true },
  { header: '付款工具',      source: 'contribution', apiField: 'payment_instrument_id',  options: true },
  { header: '姓氏',          source: 'contact',      apiField: 'last_name' },
  { header: '名字',          source: 'contact',      apiField: 'first_name' },
  { header: '收據抬頭／姓名', source: 'custom',      customKey: 'receipt_name' },
  { header: '捐款徵信名稱',  source: 'custom',       customKey: 'credit_name' },
  { header: '報稅憑證',      source: 'custom',       customKey: 'tax_cert' },
];

/** 從 FIELD_DEFS 產生 Contribution.get 所需的 return.* 參數。 */
function buildReturnFields(fieldMap) {
  const fields = { 'return.contact_id': 1 };
  FIELD_DEFS.forEach(function(d) {
    if (d.source === 'contribution') {
      fields['return.' + d.apiField] = 1;
    } else if (d.source === 'custom') {
      const key = fieldMap[d.customKey];
      if (key) fields['return.' + key] = 1;
    }
    // contact 欄位由 api.Contact.get 處理，不需加入 return.*
  });
  return fields;
}

/**
 * 預先取得所有標記 options:true 欄位的 value→label 對應表。
 * 使用 Contribution.getoptions 一次取得一個欄位的所有選項。
 *
 * @param {Object} config
 * @returns {Object}  { apiField: { '1': 'label', ... }, ... }
 */
function fetchOptionLabels(config) {
  const labels = {};
  FIELD_DEFS.forEach(function(d) {
    if (!d.options) return;
    Logger.log('Fetching options for: ' + d.apiField);
    const result = callApi(config, 'Contribution', 'getoptions', { field: d.apiField });
    const map = {};
    (result.values || []).forEach(function(opt) {
      map[String(opt.value)] = opt.label;
    });
    labels[d.apiField] = map;
  });
  return labels;
}

/**
 * 分頁抓取日期區間所有捐款（含聯絡人姓名與自訂欄位）。
 * API 單次上限 100 筆；每頁請求間隔 500ms 以符合速率限制。
 * 終止條件：offset 超過 total，或 values 回傳空陣列（已無資料）。
 *
 * @param {Object} config
 * @param {string} startDate
 * @param {string} endDate
 * @param {number} total         getcount 回傳的總筆數
 * @param {Object} optionLabels  fetchOptionLabels() 回傳的 value→label 對應表
 * @returns {Array}  每元素為以 header 名稱為 key 的物件
 */
function fetchAllContributions(config, startDate, endDate, total, optionLabels) {
  const rows     = [];
  const pageSize = config.pageSize || 100;
  let   offset   = 0;

  while (offset < total) {
    Logger.log('Fetching offset=' + offset + ' / ' + total);

    const result = callApi(config, 'Contribution', 'get', {
      ...buildDateRange(startDate, endDate),
      ...buildReturnFields(config.fieldMap),
      'api.Contact.get': {
        id: '$value.contact_id',
        'return.last_name':  1,
        'return.first_name': 1,
      },
      options: { limit: pageSize, offset: offset },
    });

    // values 為陣列（API v3 標準回應格式）
    const values = Array.isArray(result.values) ? result.values : Object.values(result.values || {});
    if (values.length === 0) break;  // 已無資料，提前結束

    values.forEach(function(c) {
      rows.push(parseRow(c, config.fieldMap, optionLabels));
    });

    offset += pageSize;

    // 速率限制：API 要求請求間隔至少 0.5s（尚有下一頁才等待）
    if (offset < total) Utilities.sleep(500);
  }

  return rows;
}

/** 從 FIELD_DEFS 派生的表頭陣列，供 ensureHeader 使用。 */
const HEADERS = FIELD_DEFS.map(function(d) { return d.header; });

/**
 * 依 FIELD_DEFS 將單筆捐款物件解析為以 header 名稱為 key 的物件。
 * @param {Object} c             API 回傳的單筆捐款物件
 * @param {Object} fieldMap      自訂欄位 key map
 * @param {Object} optionLabels  value→label 對應表（fetchOptionLabels 回傳）
 * @returns {Object}
 */
function parseRow(c, fieldMap, optionLabels) {
  const contact = ((c['api.Contact.get'] || {}).values || [])[0] || {};
  const row = {};
  FIELD_DEFS.forEach(function(d) {
    if (d.source === 'contribution') {
      const raw = c[d.apiField];
      if (d.options) {
        const labelMap = (optionLabels && optionLabels[d.apiField]) || {};
        row[d.header] = labelMap[String(raw)] || raw || '';
      } else {
        row[d.header] = d.format ? formatDate(raw) : (raw || '');
      }
    } else if (d.source === 'contact') {
      row[d.header] = contact[d.apiField] || '';
    } else if (d.source === 'custom') {
      const key = fieldMap[d.customKey];
      row[d.header] = (key && c[key] !== undefined) ? c[key] : '';
    }
  });
  return row;
}

// ─── 寫入 Google Sheet ────────────────────────────────────────────────────────

/**
 * 將捐款列 upsert 至 netiCRM捐款同步表格（以「捐款編號」欄為唯一鍵）。
 * 依 Sheet 第一列實際 header 名稱對應欄位，不依賴欄位順序。
 *
 * @param {Object} config
 * @param {Array}  rows       parseRow() 產生的物件陣列
 * @param {string} startDate  本次同步開始日（用於 log）
 * @param {string} endDate    本次同步結束日（用於 log）
 */
function writeToSheet(config, rows, startDate, endDate) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('找不到工作表：' + config.sheetName);

  // 確保表頭存在
  ensureHeader(sheet);

  // 掃描第一列，建立 header 名稱 → 欄位索引（0-based）map
  const lastCol   = sheet.getLastColumn();
  const headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const headerMap = {};
  headerRow.forEach(function(name, i) {
    if (name !== '') headerMap[String(name)] = i;
  });

  // 將 row 物件依 headerRow 順序轉為陣列（找不到的 header 填空字串）
  function rowToArray(rowObj) {
    return headerRow.map(function(name) {
      if (name === '') return '';
      return rowObj[name] !== undefined ? rowObj[name] : '';
    });
  }

  // 建立現有捐款編號 → 列號 map（依 header 定位「捐款編號」欄）
  const idColIdx    = headerMap['捐款編號'];
  const lastRow     = sheet.getLastRow();
  const existingMap = {};
  if (idColIdx !== undefined && lastRow >= 2) {
    const idCol = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
    idCol.forEach(function(r, i) {
      if (r[0] !== '') existingMap[String(r[0])] = i + 2;
    });
  }

  let inserted = 0, updated = 0;
  rows.forEach(function(row) {
    const contribId = String(row['捐款編號']);
    const rowArr    = rowToArray(row);
    if (existingMap[contribId] !== undefined) {
      sheet.getRange(existingMap[contribId], 1, 1, headerRow.length).setValues([rowArr]);
      updated++;
    } else {
      sheet.appendRow(rowArr);
      existingMap[contribId] = sheet.getLastRow();
      inserted++;
    }
  });

  Logger.log('[' + startDate + ' ~ ' + endDate + '] inserted=' + inserted + ', updated=' + updated);
}

/** 若第一列不是表頭，插入表頭列。 */
function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow[7] !== '捐款編號') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

// ─── 共用工具 ─────────────────────────────────────────────────────────────────

/**
 * 建立日期範圍過濾條件。
 * 結束日時間固定為 23:59:59，確保涵蓋當天所有捐款。
 *
 * @param {string} startDate  'yyyy-MM-dd'
 * @param {string} endDate    'yyyy-MM-dd'
 * @returns {Object}  Contribution.get / getcount 日期區間條件
 */
function buildDateRange(startDate, endDate) {
  return {
    contribution_date_low:       startDate,
    contribution_date_low_time:  '00:00:00',
    contribution_date_high:      endDate,
    contribution_date_high_time: '23:59:59',
    contribution_status_id:      [1, 3],  // 1=已完成, 3=已取消
  };
}

/** 格式化日期字串，移除多餘時間（若無值回傳空字串）。 */
function formatDate(value) {
  if (!value) return '';
  return String(value).replace('T', ' ').substring(0, 19);
}

/**
 * 呼叫 netiCRM REST API v3。
 *
 * 認證：HTTP headers（x-civicrm-api-key / x-civicrm-site-key）
 * 端點：<baseUrl>?entity=<Entity>&action=<action>
 * 請求：POST，Content-Type: application/json，body 為 JSON 字串
 * 回應：{ "is_error": 0, "version": 3, "count": N, "values": [...] }
 *       getcount 特例：{ "is_error": 0, "result": N }
 *
 * @param {Object} config
 * @param {string} entity   例如 'Contribution'
 * @param {string} action   例如 'get' | 'getcount' | 'create' | 'delete'
 * @param {Object} params   API 參數物件
 * @returns {Object}        API 回傳的 JSON 物件
 */
function callApi(config, entity, action, params) {
  const url = config.baseUrl
    + '?entity=' + encodeURIComponent(entity)
    + '&action=' + encodeURIComponent(action);

  const options = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-civicrm-api-key':  config.apiKey,
      'x-civicrm-site-key': config.siteKey,
    },
    payload:            JSON.stringify(params),
    muteHttpExceptions: true,
  };

  Logger.log('[callApi] ' + entity + '.' + action + ' params=' + JSON.stringify(params));

  const response = UrlFetchApp.fetch(url, options);
  const code     = response.getResponseCode();
  const rawText  = response.getContentText();

  if (code !== 200) {
    throw new Error('API HTTP error ' + code + ': ' + rawText.substring(0, 200));
  }

  const json = JSON.parse(rawText);
  if (json.is_error) {
    throw new Error('API error: ' + (json.error_message || JSON.stringify(json)));
  }

  Logger.log('[callApi] HTTP ' + code + ' count=' + (json.count || json.result || 0));
  const logValues = Array.isArray(json.values) ? json.values : Object.values(json.values || {});
  logValues.forEach(function(item, i) {
    Logger.log('[callApi] values[' + i + '] ' + JSON.stringify(item));
  });

  return json;
}
