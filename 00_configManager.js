/**
 * 設定値管理モジュール（シンプル版）
 * 環境変数のような感覚で設定値を管理
 */

// ===== 設定シート名 =====
var CONFIG_SHEET_NAME = "設定";

// ===== 設定値キャッシュ =====
var _config = null;

const MAX_OUTPUT_ITEMS = 20; // C〜V（20列）に合わせる場合
const COL_C = 3; // C列（出力起点）
const APPLY_CLIP_WRAP = true; // ← 文字折返しで高さ変化を防ぐなら true
const MAX_OUTPUT_COLS = 20; // C～V

// ===== 設定値読み取り関数 =====

/**
 * 設定値を読み込み（シンプル版）
 */
function loadConfig_() {
  if (_config) return _config; // 一度読み込んだらキャッシュを使用

  try {
    var sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
    if (!sheet) {
      console.warn("設定シートが見つかりません、デフォルト値を使用");
      _config = getDefaultConfig_();
      return _config;
    }

    var config = {};
    var lastRow = sheet.getLastRow();

    // A列からキー、B列から値を読み取り（シンプル構造）
    for (var row = 2; row <= lastRow; row++) {
      var key = sheet.getRange(row, 1).getValue();
      var value = sheet.getRange(row, 2).getValue();

      if (key) {
        config[key] = value;
      }
    }

    _config = config;
    console.log("設定値を読み込みました:", Object.keys(config).length + "項目");
    return config;
  } catch (e) {
    console.error("設定値読み込みエラー:", e.message);
    _config = getDefaultConfig_();
    return _config;
  }
}

/**
 * デフォルト設定値（環境変数スタイル）
 */
function getDefaultConfig_() {
  return {
    // SBA設定
    SBA_NAME: "SBA",
    SBA_URL_ROW: 11,
    SBA_URL_COL: 2,
    SBA_START_ROW: 5,
    SBA_END_ROW: 16,

    // エコリング設定
    ECORING_NAME: "エコリング",
    ECORING_URL_ROW: 17,
    ECORING_URL_COL: 2,
    ECORING_START_ROW: 18,
    ECORING_END_ROW: 29,

    // 楽天設定
    RAKUTEN_NAME: "楽天",
    RAKUTEN_URL_ROW: 30,
    RAKUTEN_URL_COL: 2,
    RAKUTEN_START_ROW: 31,
    RAKUTEN_END_ROW: 42,

    // オークファン設定
    AUCFAN_NAME: "オークファン",
    AUCFAN_URL_ROW: 44,
    AUCFAN_URL_COL: 2,
    AUCFAN_START_ROW: 45,
    AUCFAN_END_ROW: 53,

    // ヤフオク設定
    YAHOO_NAME: "ヤフオク",
    YAHOO_URL_ROW: 44,
    YAHOO_URL_COL: 2,
    YAHOO_START_ROW: 57,
    YAHOO_END_ROW: 66,

    // 出力設定
    OUTPUT_MAX_ITEMS: 20,
    OUTPUT_MAX_COLS: 20,
    OUTPUT_START_COL: 3,
    OUTPUT_CLIP_WRAP: true,

    // 行オフセット
    OFFSET_DETAIL_URL: 0,
    OFFSET_IMAGE_URL: 1,
    OFFSET_PRODUCT_NAME: 4,
    OFFSET_DATE: 5,
    OFFSET_RANK: 6,
    OFFSET_SOURCE: 9,
    OFFSET_PRICE: 12,
    OFFSET_SHOP: 13,
  };
}

/**
 * 設定値キャッシュをクリア
 */
function clearConfigCache_() {
  _config = null;
}

// ===== シンプルな設定値取得関数 =====

/**
 * 設定値を取得（環境変数スタイル）
 */
function getConfig_(key) {
  var config = loadConfig_();
  return config[key];
}

/**
 * アクティブなシートを取得
 */
function getDataCollectionSheet_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) {
    throw new Error("アクティブなシートが見つかりません");
  }
  return sheet;
}

/**
 * 出力設定を取得
 */
function getOutputSettings_() {
  return {
    maxItems: getConfig_("OUTPUT_MAX_ITEMS") || 20,
    maxCols: getConfig_("OUTPUT_MAX_COLS") || 20,
    startCol: getConfig_("OUTPUT_START_COL") || 3,
    applyClipWrap: getConfig_("OUTPUT_CLIP_WRAP") !== false,
  };
}

/**
 * サイト設定を取得
 */
function getSiteConfig_(siteKey) {
  var upperSite = siteKey.toUpperCase();
  return {
    name: getConfig_(upperSite + "_NAME"),
    urlRow: getConfig_(upperSite + "_URL_ROW"),
    urlCol: getConfig_(upperSite + "_URL_COL"),
    startRow: getConfig_(upperSite + "_START_ROW"),
    endRow: getConfig_(upperSite + "_END_ROW"),
  };
}

/**
 * サイトのURL位置を取得
 */
function getSiteUrlPosition_(siteKey) {
  var config = getSiteConfig_(siteKey);
  return {
    row: config.urlRow,
    col: config.urlCol,
  };
}

/**
 * サイトの書き込み範囲を取得
 */
function getSiteWriteRange_(siteKey) {
  var config = getSiteConfig_(siteKey);
  return {
    startRow: config.startRow,
    endRow: config.endRow,
    name: config.name,
  };
}

// ===== 設定シート作成関数 =====

/**
 * 設定シートを作成（シンプル版）
 */
function createConfigSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
    }

    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, 3).setValues([["キー", "値", "説明"]]);

    // ヘッダー行のスタイル設定
    var headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("white");
    headerRange.setFontWeight("bold");

    // 設定データを準備（環境変数スタイル）
    var configData = [
      // SBA設定
      ["SBA_NAME", "SBA", "SBAの表示名"],
      ["SBA_URL_ROW", 4, "SBAのURL行番号"],
      ["SBA_URL_COL", 2, "SBAのURL列番号"],
      ["SBA_START_ROW", 12, "SBAデータ開始行"],
      ["SBA_END_ROW", 25, "SBAデータ終了行"],

      // エコリング設定
      ["ECORING_NAME", "エコリング", "エコリングの表示名"],
      ["ECORING_URL_ROW", 17, "エコリングのURL行番号"],
      ["ECORING_URL_COL", 2, "エコリングのURL列番号"],
      ["ECORING_START_ROW", 18, "エコリングデータ開始行"],
      ["ECORING_END_ROW", 29, "エコリングデータ終了行"],

      // 楽天設定
      ["RAKUTEN_NAME", "楽天", "楽天の表示名"],
      ["RAKUTEN_URL_ROW", 30, "楽天のURL行番号"],
      ["RAKUTEN_URL_COL", 2, "楽天のURL列番号"],
      ["RAKUTEN_START_ROW", 31, "楽天データ開始行"],
      ["RAKUTEN_END_ROW", 42, "楽天データ終了行"],

      // オークファン設定
      ["AUCFAN_NAME", "オークファン", "オークファンの表示名"],
      ["AUCFAN_URL_ROW", 44, "オークファンのURL行番号"],
      ["AUCFAN_URL_COL", 2, "オークファンのURL列番号"],
      ["AUCFAN_START_ROW", 45, "オークファンデータ開始行"],
      ["AUCFAN_END_ROW", 53, "オークファンデータ終了行"],

      // ヤフオク設定
      ["YAHOO_NAME", "ヤフオク", "ヤフオクの表示名"],
      ["YAHOO_URL_ROW", 44, "ヤフオクのURL行番号"],
      ["YAHOO_URL_COL", 2, "ヤフオクのURL列番号"],
      ["YAHOO_START_ROW", 57, "ヤフオクデータ開始行"],
      ["YAHOO_END_ROW", 66, "ヤフオクデータ終了行"],

      // 出力設定
      ["OUTPUT_MAX_ITEMS", 20, "最大出力アイテム数"],
      ["OUTPUT_MAX_COLS", 20, "最大出力列数"],
      ["OUTPUT_START_COL", 3, "出力開始列（C列=3）"],
      ["OUTPUT_CLIP_WRAP", true, "折り返し抑制を適用するか"],

      // 行オフセット設定
      ["OFFSET_DETAIL_URL", 0, "詳細URL行オフセット"],
      ["OFFSET_IMAGE_URL", 1, "画像URL行オフセット"],
      ["OFFSET_PRODUCT_NAME", 4, "商品名行オフセット"],
      ["OFFSET_DATE", 5, "日付行オフセット"],
      ["OFFSET_RANK", 6, "ランク行オフセット"],
      ["OFFSET_SOURCE", 9, "引用サイト行オフセット"],
      ["OFFSET_PRICE", 12, "価格行オフセット"],
      ["OFFSET_SHOP", 13, "ショップ行オフセット"],
    ];

    // データを書き込み
    if (configData.length > 0) {
      sheet.getRange(2, 1, configData.length, 3).setValues(configData);
    }

    // 列幅を調整
    sheet.setColumnWidth(1, 200); // キー列
    sheet.setColumnWidth(2, 100); // 値列
    sheet.setColumnWidth(3, 300); // 説明列

    // 設定値キャッシュをクリア
    clearConfigCache_();

    console.log("設定シートを作成しました");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "設定シートを作成しました",
      "✅ 完了",
      3
    );
  } catch (e) {
    console.error("設定シート作成エラー:", e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "エラー: " + e.message,
      "❌ エラー",
      5
    );
  }
}

/**
 * 設定値をリロード
 */
function reloadConfig() {
  clearConfigCache_();
  var config = loadConfig_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "設定値をリロードしました",
    "✅ 完了",
    2
  );
  console.log("設定値をリロードしました:", Object.keys(config).length + "項目");
}

/**
 * 設定シートの状況を確認
 */
function checkConfigSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    if (!sheet) {
      console.log("設定シートが存在しません");
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "設定シートが存在しません。createConfigSheet()を実行してください。",
        "⚠️ 警告",
        5
      );
      return false;
    }

    // 設定値を確認
    clearConfigCache_();
    var config = loadConfig_();

    console.log("現在の設定値:", {
      SBA_URL_ROW: config.SBA_URL_ROW,
      ECORING_URL_ROW: config.ECORING_URL_ROW,
      RAKUTEN_URL_ROW: config.RAKUTEN_URL_ROW,
      AUCFAN_URL_ROW: config.AUCFAN_URL_ROW,
      YAHOO_URL_ROW: config.YAHOO_URL_ROW,
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "設定シート確認完了。コンソールログを確認してください。",
      "✅ 完了",
      3
    );

    return true;
  } catch (e) {
    console.error("設定シート確認エラー:", e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "エラー: " + e.message,
      "❌ エラー",
      5
    );
    return false;
  }
}
