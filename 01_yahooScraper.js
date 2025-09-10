/**
 * ヤフオク専用スクレイパー
 * ヤフオクからの商品情報取得処理
 */

// ===== ヤフオク スクレイピング =====

/**
 * ヤフオクのHTMLを取得
 */
function fetchYahooAuctionHtml_(url) {
    try {
      logInfo_("fetchYahooAuctionHtml_", "ヤフオクHTML取得開始", "URL: " + url);
  
      // URLが古い形式の場合は新しい形式に変換
      var updatedUrl = convertToNewYahooUrl_(url);
      if (updatedUrl !== url) {
        logInfo_(
          "fetchYahooAuctionHtml_",
          "URL変換",
          "旧URL: " + url + " → 新URL: " + updatedUrl
        );
        url = updatedUrl;
      }
  
      var options = getCommonHttpOptions_();
      var response = UrlFetchApp.fetch(url, options);
  
      validateHttpResponse_(response, "ヤフオク");
      return getResponseTextWithBestCharset_(response);
    } catch (e) {
      logError_(
        "fetchYahooAuctionHtml_",
        "ERROR",
        "HTMLフェッチエラー",
        e.message
      );
      throw new Error("ヤフオク URLからのHTMLフェッチエラー: " + e.message);
    }
  }
  
  /**
   * 古いヤフオクURLを新しい形式に変換
   */
  function convertToNewYahooUrl_(url) {
    // 古い closedsearch 形式を新しい search 形式に変換
    if (url.includes("/closedsearch/closedsearch")) {
      // URLパラメータを解析
      var urlParts = url.split("?");
      if (urlParts.length > 1) {
        var paramString = urlParts[1];
        var keyword = "";
        var category = "";
  
        // パラメータを手動で解析
        var paramPairs = paramString.split("&");
        for (var i = 0; i < paramPairs.length; i++) {
          var pair = paramPairs[i].split("=");
          if (pair.length === 2) {
            var key = pair[0];
            var value = decodeURIComponent(pair[1]);
  
            if (key === "p") {
              keyword = value;
            } else if (key === "auccat") {
              category = value;
            }
          }
        }
  
        // 新しいURL形式で構築
        var newUrl = "https://auctions.yahoo.co.jp/search/search";
        var newParams = [];
  
        if (keyword) {
          newParams.push("p=" + encodeURIComponent(keyword));
        }
        if (category) {
          newParams.push("auccat=" + category);
        }
  
        // 落札済み商品を検索するパラメータを追加
        newParams.push("s1=end"); // 終了したオークション
        newParams.push("o1=a"); // 価格順（昇順）
        newParams.push("mode=2"); // 落札相場モード
  
        if (newParams.length > 0) {
          newUrl += "?" + newParams.join("&");
        }
  
        return newUrl;
      }
    }
  
    return url;
  }
  
  /**
   * ヤフオクのHTMLをパースして商品情報を抽出
   */
  function parseYahooAuctionFromHtml_(html) {
    var items = [];
    var H = String(html);
  
    console.log("ヤフオクパーサー - HTMLの長さ:", H.length);
  
    // 複数のパターンで商品ブロックを抽出を試行
    var blocks = extractYahooProductBlocks_(H);
  
    if (!blocks || blocks.length === 0) {
      console.log("ヤフオク 商品ブロックが見つかりません");
      // デバッグ用：HTMLの一部を出力
      console.log("HTML先頭500文字:", H.substring(0, 500));
      return items;
    }
  
    console.log("ヤフオク 商品ブロック数:", blocks.length);
  
    blocks.forEach(function (block, index) {
      var item = parseYahooProductBlock_(block, index);
      if (item) {
        items.push(item);
      }
    });
  
    console.log("ヤフオク パース完了:", items.length + "件");
    return items;
  }
  
  /**
   * ヤフオクの商品ブロックを抽出（複数パターンに対応）
   */
  function extractYahooProductBlocks_(html) {
    var blocks = [];
  
    // パターン1: 古い Product クラス
    var pattern1 = /<li\s+class="Product"[^>]*>/gi;
    var positions1 = [];
    var match;
  
    while ((match = pattern1.exec(html)) !== null) {
      positions1.push(match.index);
    }
  
    if (positions1.length > 0) {
      console.log(
        "パターン1（Product）で",
        positions1.length,
        "件見つかりました"
      );
      var maxProducts = Math.min(50, positions1.length);
      for (var i = 0; i < maxProducts; i++) {
        var start = positions1[i];
        var end = i + 1 < positions1.length ? positions1[i + 1] : start + 10000;
        blocks.push(html.substring(start, end));
      }
      return blocks;
    }
  
    // パターン2: 新しい構造（data-auction-id属性を持つ要素）
    var pattern2 = /<[^>]+data-auction-id="[^"]*"[^>]*>/gi;
    var positions2 = [];
  
    while ((match = pattern2.exec(html)) !== null) {
      positions2.push(match.index);
    }
  
    if (positions2.length > 0) {
      console.log(
        "パターン2（data-auction-id）で",
        positions2.length,
        "件見つかりました"
      );
      var maxProducts = Math.min(50, positions2.length);
      for (var i = 0; i < maxProducts; i++) {
        var start = positions2[i];
        var end = i + 1 < positions2.length ? positions2[i + 1] : start + 8000;
        blocks.push(html.substring(start, end));
      }
      return blocks;
    }
  
    // パターン3: 一般的なリストアイテム
    var pattern3 = /<li[^>]*class="[^"]*item[^"]*"[^>]*>/gi;
    var positions3 = [];
  
    while ((match = pattern3.exec(html)) !== null) {
      positions3.push(match.index);
    }
  
    if (positions3.length > 0) {
      console.log("パターン3（item）で", positions3.length, "件見つかりました");
      var maxProducts = Math.min(50, positions3.length);
      for (var i = 0; i < maxProducts; i++) {
        var start = positions3[i];
        var end = i + 1 < positions3.length ? positions3[i + 1] : start + 8000;
        blocks.push(html.substring(start, end));
      }
      return blocks;
    }
  
    return blocks;
  }
  
  /**
   * 個別の商品ブロックをパース
   */
  function parseYahooProductBlock_(block, index) {
    var title = "";
    var detailUrl = "";
    var imageUrl = "";
    var price = "";
    var date = "";
    var shop = "";
  
    // 商品名とURL（複数パターンで試行）
    var titlePatterns = [
      /<a[^>]*class="Product__titleLink"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i,
      /<a[^>]*href="([^"]*\/[a-z]\d+)"[^>]*[^>]*title="([^"]*)"[^>]*>/i,
      /<a[^>]*href="([^"]*\/[a-z]\d+)"[^>]*>([\s\S]*?)<\/a>/i,
      /<a[^>]*href="([^"]+)"[^>]*class="[^"]*title[^"]*"[^>]*>([\s\S]*?)<\/a>/i,
    ];
  
    for (var i = 0; i < titlePatterns.length; i++) {
      var titleMatch = block.match(titlePatterns[i]);
      if (titleMatch) {
        detailUrl = titleMatch[1];
        var titleText = titleMatch[2];
        titleText = titleText.replace(/<[^>]+>/g, "");
        titleText = titleText.replace(/\s+/g, " ");
        title = htmlDecode_(titleText).trim();
  
        if (detailUrl && !detailUrl.startsWith("http")) {
          detailUrl = "https://auctions.yahoo.co.jp" + detailUrl;
        }
        break;
      }
    }
  
    // 画像URL（複数パターンで試行）
    var imagePatterns = [
      /<img[^>]*class="Product__imageData"[^>]*src="([^"]+)"/i,
      /<img[^>]*src="([^"]+)"[^>]*alt="[^"]*"/i,
      /<img[^>]*src="([^"]+)"/i,
    ];
  
    for (var i = 0; i < imagePatterns.length; i++) {
      var imgMatch = block.match(imagePatterns[i]);
      if (imgMatch) {
        imageUrl = htmlDecode_(imgMatch[1]);
        break;
      }
    }
  
    // 落札価格（複数パターンで試行）
    var pricePatterns = [
      /<span[^>]*class="Product__label"[^>]*>落札<\/span>[\s\S]*?<span[^>]*class="Product__priceValue[^"]*"[^>]*>([^<]+)<\/span>/i,
      /<span[^>]*class="[^"]*price[^"]*"[^>]*>([^<]*円[^<]*)<\/span>/i,
      /落札価格[^>]*>([^<]*\d+[^<]*円[^<]*)</i,
      /(\d+(?:,\d{3})*)\s*円/i,
    ];
  
    for (var i = 0; i < pricePatterns.length; i++) {
      var priceMatch = block.match(pricePatterns[i]);
      if (priceMatch) {
        var priceText = htmlDecode_(priceMatch[1]).trim();
        price = normalizeNumber_(priceText.replace(/円/g, ""));
        break;
      }
    }
  
    // 終了日時（複数パターンで試行）
    var datePatterns = [
      /<span[^>]*class="Product__time"[^>]*>([^<]+)<\/span>/i,
      /<span[^>]*class="[^"]*time[^"]*"[^>]*>([^<]+)<\/span>/i,
      /終了日時[^>]*>([^<]+)</i,
      /(\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2})/i,
    ];
  
    for (var i = 0; i < datePatterns.length; i++) {
      var dateMatch = block.match(datePatterns[i]);
      if (dateMatch) {
        date = htmlDecode_(dateMatch[1]).trim();
        break;
      }
    }
  
    // ショップ情報
    if (block.indexOf("年間ベストストア") > -1) {
      shop = "ベストストア";
    } else if (block.indexOf("ストア") > -1) {
      shop = "ストア";
    }
  
    // デバッグ情報（最初の3商品）
    if (index < 3) {
      console.log("=== ヤフオク商品 " + (index + 1) + " ===");
      console.log("商品名:", title);
      console.log("価格:", price);
      console.log("日時:", date);
    }
  
    if (title && (detailUrl || price || date)) {
      return createItemData_({
        title: title,
        detailUrl: detailUrl,
        imageUrl: imageUrl,
        date: date,
        rank: "", // ヤフオクにはランク情報なし
        price: price,
        shop: shop,
        source: "ヤフオク",
        soldOut: false,
      });
    }
  
    return null;
  }
  
  /**
   * ヤフオクの修正をテストする関数
   */
  function testYahooScraperFix() {
    try {
      console.log("=== ヤフオクスクレイパー修正テスト開始 ===");
  
      // 古いURL形式をテスト
      var oldUrl =
        "https://auctions.yahoo.co.jp/closedsearch/closedsearch?auccat=2&tab_ex=commerce&ei=utf-8&aq=-1&oq=&sc_i=&p=%E3%82%BB%E3%83%AA%E3%83%BC%E3%83%8C%20%E3%83%B4%E3%82%A3%E3%82%AF%E3%83%88%E3%83%AF%E3%83%BC%E3%83%AB%20";
      console.log("テスト用の古いURL:", oldUrl);
  
      // URL変換をテスト
      var newUrl = convertToNewYahooUrl_(oldUrl);
      console.log("変換後の新しいURL:", newUrl);
  
      // HTMLを取得してテスト
      console.log("HTMLフェッチテスト開始...");
      var html = fetchYahooAuctionHtml_(oldUrl);
      console.log("HTMLフェッチ成功。長さ:", html.length);
  
      // パースをテスト
      console.log("HTMLパーステスト開始...");
      var items = parseYahooAuctionFromHtml_(html);
      console.log("パース完了。取得件数:", items.length);
  
      // 結果を表示
      if (items.length > 0) {
        console.log("=== 取得した商品例 ===");
        for (var i = 0; i < Math.min(3, items.length); i++) {
          console.log("商品" + (i + 1) + ":");
          console.log("  タイトル:", items[i].title);
          console.log("  価格:", items[i].price);
          console.log("  日付:", items[i].date);
        }
      }
  
      console.log("=== ヤフオクスクレイパー修正テスト完了 ===");
      return items;
    } catch (e) {
      console.error("ヤフオクスクレイパーテストエラー:", e.message);
      console.error("スタックトレース:", e.stack);
      throw e;
    }
  }
  
  /**
   * ヤフオクの出力位置を修正するための設定更新関数
   */
  function updateYahooOutputPosition() {
    try {
      console.log("=== ヤフオク出力位置修正開始 ===");
  
      // 設定値キャッシュをクリア
      clearConfigCache_();
  
      // 新しい設定値を読み込み
      var config = loadConfig_();
  
      console.log("現在のヤフオク設定:");
      console.log("  URL行:", config.YAHOO_URL_ROW);
      console.log("  開始行:", config.YAHOO_START_ROW);
      console.log("  終了行:", config.YAHOO_END_ROW);
  
      // 設定シートが存在するか確認
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
      if (sheet) {
        console.log("設定シートが見つかりました");
  
        // 設定値を確認
        var yahooStartRow = null;
        var lastRow = sheet.getLastRow();
  
        for (var row = 2; row <= lastRow; row++) {
          var key = sheet.getRange(row, 1).getValue();
          if (key === "YAHOO_START_ROW") {
            yahooStartRow = sheet.getRange(row, 2).getValue();
            break;
          }
        }
  
        console.log("設定シートのYAHOO_START_ROW:", yahooStartRow);
  
        if (yahooStartRow === 57) {
          console.log("✅ 設定値は正しく57に設定されています");
        } else {
          console.log(
            "⚠️ 設定値が57ではありません。設定シートを確認してください。"
          );
        }
      } else {
        console.log(
          "⚠️ 設定シートが見つかりません。createConfigSheet()を実行してください。"
        );
      }
  
      console.log("=== ヤフオク出力位置修正完了 ===");
  
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "ヤフオクの出力位置を57行目に修正しました",
        "✅ 完了",
        3
      );
    } catch (e) {
      console.error("ヤフオク出力位置修正エラー:", e.message);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "エラー: " + e.message,
        "❌ エラー",
        5
      );
    }
  }
  