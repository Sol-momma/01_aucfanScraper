/**
 * スプレッドシートボタン用関数
 * 各サイトのデータ取得・クリア処理をボタンから実行するための関数群
 */

// ===== 個別サイト データ取得関数 =====

/**
 * SBAデータ取得ボタン用関数
 */
function buttonGetSbaData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("データ収集シートが見つかりません");
        return;
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "SBAデータを取得中...",
        "処理中",
        3
      );
  
      // データ取得前に既存データをクリア
      clearSectionForDataFetch_(sheet, "SBA");
  
      var sbaItems = getSbaData_(sheet);
  
      if (sbaItems && sbaItems.length > 0) {
        writeItemsToSpecificSection_(
          sheet,
          sbaItems,
          5, // SBAセクションの開始行（詳細URLの行）
          "SBA"
        );
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `SBAデータを${sbaItems.length}件取得しました`,
          "✅ 完了",
          3
        );
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "SBAデータが取得できませんでした",
          "⚠️ 警告",
          3
        );
      }
    } catch (e) {
      console.error("SBAデータ取得エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  
  /**
   * エコリングデータ取得ボタン用関数
   */
  function buttonGetEcoringData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エコリングデータを取得中...",
        "処理中",
        3
      );
  
      // データ取得前に既存データをクリア
      clearSectionForDataFetch_(sheet, "エコリング");
  
      var ecoringItems = getEcoringData_(sheet);
  
      if (ecoringItems && ecoringItems.length > 0) {
        writeItemsToSpecificSection_(sheet, ecoringItems, 18, "エコリング");
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `エコリングデータを${ecoringItems.length}件取得しました`,
          "✅ 完了",
          3
        );
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "エコリングデータが取得できませんでした",
          "⚠️ 警告",
          3
        );
      }
    } catch (e) {
      console.error("エコリングデータ取得エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  
  /**
   * 楽天データ取得ボタン用関数
   */
  function buttonGetRakutenData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "楽天データを取得中...",
        "処理中",
        3
      );
  
      // データ取得前に既存データをクリア
      clearSectionForDataFetch_(sheet, "楽天");
  
      var rakutenItems = getRakutenData_(sheet);
  
      if (rakutenItems && rakutenItems.length > 0) {
        writeItemsToSpecificSection_(sheet, rakutenItems, 31, "楽天");
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `楽天データを${rakutenItems.length}件取得しました`,
          "✅ 完了",
          3
        );
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "楽天データが取得できませんでした",
          "⚠️ 警告",
          3
        );
      }
    } catch (e) {
      console.error("楽天データ取得エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  
  /**
   * オークファンデータ取得ボタン用関数（別シート対応）
   */
  function buttonGetAucfanData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("オークファンシートが見つかりません");
        return;
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "オークファンデータを取得中...",
        "処理中",
        3
      );
  
      // データ取得前に既存データをクリア（別シート専用）
      clearAucfanSeparateSheet_(sheet);
  
      // 別シート専用のデータ取得
      var aucfanItems = getAucfanDataFromSeparateSheet_(sheet);
  
      if (aucfanItems && aucfanItems.length > 0) {
        // 別シート専用のデータライター
        writeAucfanDataToSeparateSheet_(sheet, aucfanItems);
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `オークファンデータを${aucfanItems.length}件取得しました`,
          "✅ 完了",
          3
        );
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "オークファンデータが取得できませんでした\n15行目以降にHTMLを貼り付けてください",
          "⚠️ 警告",
          5
        );
      }
    } catch (e) {
      console.error("オークファンデータ取得エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  
  /**
   * ヤフオクデータ取得ボタン用関数
   */
  function buttonGetYahooData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "ヤフオクデータを取得中...",
        "処理中",
        3
      );
  
      // データ取得前に既存データをクリア
      var yahooConfig = getSiteWriteRange_("yahoo");
      clearSectionForDataFetch_(sheet, "ヤフオク");
  
      var yahooItems = getYahooAuctionData_(sheet);
  
      if (yahooItems && yahooItems.length > 0) {
        var yahooConfig = getSiteWriteRange_("yahoo");
        writeItemsToSpecificSection_(
          sheet,
          yahooItems,
          yahooConfig.startRow,
          "ヤフオク"
        );
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `ヤフオクデータを${yahooItems.length}件取得しました`,
          "✅ 完了",
          3
        );
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "ヤフオクデータが取得できませんでした",
          "⚠️ 警告",
          3
        );
      }
    } catch (e) {
      console.error("ヤフオクデータ取得エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  
  // ===== 個別サイト データクリア関数 =====
  
  /**
   * SBAデータクリアボタン用関数
   */
  function buttonClearSbaData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      clearSectionForDataFetch_(sheet, "SBA"); // 画像行をスキップしてクリア
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "SBAデータをクリアしました",
        "✅ 完了",
        2
      );
    } catch (e) {
      console.error("SBAデータクリアエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        3
      );
    }
  }
  
  /**
   * エコリングデータクリアボタン用関数
   */
  function buttonClearEcoringData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      clearSectionForDataFetch_(sheet, "エコリング"); // 画像行をスキップしてクリア
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エコリングデータをクリアしました",
        "✅ 完了",
        2
      );
    } catch (e) {
      console.error("エコリングデータクリアエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        3
      );
    }
  }
  
  /**
   * 楽天データクリアボタン用関数
   */
  function buttonClearRakutenData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      clearSectionForDataFetch_(sheet, "楽天"); // 画像行をスキップしてクリア
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "楽天データをクリアしました",
        "✅ 完了",
        2
      );
    } catch (e) {
      console.error("楽天データクリアエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        3
      );
    }
  }
  
  /**
   * オークファンデータクリアボタン用関数（別シート対応）
   */
  function buttonClearAucfanData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("オークファンシートが見つかりません");
        return;
      }
  
      // 別シート専用のクリア処理
      clearAucfanSeparateSheet_(sheet);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "オークファンデータをクリアしました",
        "✅ 完了",
        2
      );
    } catch (e) {
      console.error("オークファンデータクリアエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        3
      );
    }
  }
  
  /**
   * ヤフオクデータクリアボタン用関数
   */
  function buttonClearYahooData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      var yahooConfig = getSiteWriteRange_("yahoo");
      clearSectionForDataFetch_(sheet, "ヤフオク"); // 画像行をスキップしてクリア
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "ヤフオクデータをクリアしました",
        "✅ 完了",
        2
      );
    } catch (e) {
      console.error("ヤフオクデータクリアエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        3
      );
    }
  }
  
  // ===== 全サイト一括処理関数 =====
  
  /**
   * 全サイトデータ取得ボタン用関数
   */
  function buttonGetAllSitesData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "全サイトのデータを取得中...",
        "処理中",
        5
      );
  
      // 全サイトのデータを取得前に一括クリア（超高速）
      clearAllSitesDataFast_(sheet);
  
      var totalItems = 0;
      var results = [];
  
      // 各サイトのデータを順次取得
      try {
        var sbaItems = getSbaData_(sheet);
        if (sbaItems && sbaItems.length > 0) {
          writeItemsToSpecificSection_(sheet, sbaItems, 5, "SBA");
          totalItems += sbaItems.length;
          results.push(`SBA: ${sbaItems.length}件`);
        }
      } catch (e) {
        results.push("SBA: エラー");
        console.error("SBA取得エラー:", e.message);
      }
  
      try {
        var ecoringItems = getEcoringData_(sheet);
        if (ecoringItems && ecoringItems.length > 0) {
          writeItemsToSpecificSection_(sheet, ecoringItems, 18, "エコリング");
          totalItems += ecoringItems.length;
          results.push(`エコリング: ${ecoringItems.length}件`);
        }
      } catch (e) {
        results.push("エコリング: エラー");
        console.error("エコリング取得エラー:", e.message);
      }
  
      try {
        var rakutenItems = getRakutenData_(sheet);
        if (rakutenItems && rakutenItems.length > 0) {
          writeItemsToSpecificSection_(sheet, rakutenItems, 31, "楽天");
          totalItems += rakutenItems.length;
          results.push(`楽天: ${rakutenItems.length}件`);
        }
      } catch (e) {
        results.push("楽天: エラー");
        console.error("楽天取得エラー:", e.message);
      }
  
      // オークファンは別シートで管理するため、全サイト一括処理から除外
      // 個別のオークファンシートでbuttonGetAucfanData()を使用してください
      console.log(
        "オークファンは別シートで管理されるため、全サイト一括処理から除外されました"
      );
  
      try {
        var yahooItems = getYahooAuctionData_(sheet);
        if (yahooItems && yahooItems.length > 0) {
          var yahooConfig = getSiteWriteRange_("yahoo");
          writeItemsToSpecificSection_(
            sheet,
            yahooItems,
            yahooConfig.startRow,
            "ヤフオク"
          );
          totalItems += yahooItems.length;
          results.push(`ヤフオク: ${yahooItems.length}件`);
        }
      } catch (e) {
        results.push("ヤフオク: エラー");
        console.error("ヤフオク取得エラー:", e.message);
      }
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `全サイトデータ取得完了\n合計: ${totalItems}件\n${results.join(", ")}`,
        "✅ 完了",
        5
      );
    } catch (e) {
      console.error("全サイトデータ取得エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  
  /**
   * 全サイトデータクリアボタン用関数
   */
  function buttonClearAllSitesData() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("「データ収集」シートが見つかりません");
        return;
      }
  
      // 確認ダイアログ
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        "確認",
        "全サイトのデータをクリアしますか？\n（注意：オークファンは別シートで管理されるため除外されます）",
        ui.ButtonSet.YES_NO
      );
  
      if (response == ui.Button.YES) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "全サイトデータをクリア中...",
          "処理中",
          3
        );
  
        // 全サイトを一括でクリア（超高速）
        clearAllSitesDataFast_(sheet);
  
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "全サイトデータをクリアしました\n（オークファンは別シートで個別にクリアしてください）",
          "✅ 完了",
          5
        );
      }
    } catch (e) {
      console.error("全サイトデータクリアエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        3
      );
    }
  }
  
  // ===== オークファン複数シート対応関数 =====
  
  /**
   * オークファンテスト関数（現在のシートでHTMLパース確認）
   */
  function testAucfanCurrentSheet() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      if (!sheet) {
        SpreadsheetApp.getUi().alert("シートが見つかりません");
        return;
      }
  
      console.log("=== オークファン現在シートテスト ===");
      console.log("シート名:", sheet.getName());
  
      // 15行目以降のHTMLを確認
      var htmlFromSheet = readHtmlFromRow_(sheet, 15);
  
      if (!htmlFromSheet) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "15行目以降にHTMLが見つかりません\nB15行目以降にオークファンのHTMLを貼り付けてください",
          "⚠️ 警告",
          5
        );
        return;
      }
  
      console.log("HTMLを読み取りました:", htmlFromSheet.length + "文字");
  
      // ソース判定
      var source = detectSource_(htmlFromSheet);
      console.log("判定されたソース:", source);
  
      if (source !== "オークファン") {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "オークファンのHTMLではない可能性があります\n判定されたソース: " +
            source,
          "⚠️ 警告",
          5
        );
        return;
      }
  
      // パース実行
      var items = parseAucfanFromHtml_(htmlFromSheet);
      console.log("パース完了:", items.length + "件の商品を取得");
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `オークファンHTMLテスト完了\n取得した商品数: ${items.length}件\n詳細はログを確認してください`,
        "✅ 完了",
        5
      );
  
      // 最初の3件をログに出力
      for (var i = 0; i < Math.min(3, items.length); i++) {
        console.log("\n商品" + (i + 1) + ":");
        console.log(
          "  タイトル:",
          items[i].title ? items[i].title.substring(0, 50) + "..." : "なし"
        );
        console.log("  価格:", items[i].price || "なし");
        console.log("  日付:", items[i].date || "なし");
      }
    } catch (e) {
      console.error("オークファンテストエラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "テストエラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  