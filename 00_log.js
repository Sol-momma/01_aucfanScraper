// エラーログシートを取得または作成
function getOrCreateErrorLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errorSheet = ss.getSheetByName("エラーログ");

  if (!errorSheet) {
    // 現在のアクティブシートを記憶
    var currentSheet = ss.getActiveSheet();

    // エラーログシートが存在しない場合は作成
    errorSheet = ss.insertSheet("エラーログ");

    // ヘッダーを設定
    var headers = [
      "タイムスタンプ",
      "関数名",
      "エラータイプ",
      "エラーメッセージ",
      "詳細情報",
    ];
    errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダーのスタイルを設定
    errorSheet
      .getRange(1, 1, 1, headers.length)
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");

    // 列幅を調整
    errorSheet.setColumnWidth(1, 150); // タイムスタンプ
    errorSheet.setColumnWidth(2, 200); // 関数名
    errorSheet.setColumnWidth(3, 100); // エラータイプ
    errorSheet.setColumnWidth(4, 300); // エラーメッセージ
    errorSheet.setColumnWidth(5, 400); // 詳細情報

    // 元のシートに戻す
    if (currentSheet) {
      ss.setActiveSheet(currentSheet);
    }
  }

  return errorSheet;
}

// エラーログを記録
function logError_(functionName, errorType, errorMessage, details) {
  try {
    var errorSheet = getOrCreateErrorLogSheet_();
    var lastRow = errorSheet.getLastRow();
    var timestamp = new Date();

    // エラー情報を配列にまとめる
    var errorData = [
      timestamp,
      functionName || "不明",
      errorType || "ERROR",
      errorMessage || "エラーメッセージなし",
      details || "",
    ];

    // エラーログシートに記録
    errorSheet
      .getRange(lastRow + 1, 1, 1, errorData.length)
      .setValues([errorData]);

    // コンソールにも出力
    console.error("エラーログ:", {
      timestamp: timestamp,
      functionName: functionName,
      errorType: errorType,
      errorMessage: errorMessage,
      details: details,
    });
  } catch (logError) {
    // ログ記録自体が失敗した場合はコンソールに出力
    console.error("エラーログの記録に失敗:", logError);
  }
}

// 情報ログを記録（デバッグ用）
function logInfo_(functionName, message, details) {
  logError_(functionName, "INFO", message, details);
}

// 警告ログを記録
function logWarning_(functionName, message, details) {
  logError_(functionName, "WARNING", message, details);
}
