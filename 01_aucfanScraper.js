/**
 * オークファン専用スクレイパー
 * オークファンからの商品情報取得処理
 */

// ===== オークファン スクレイピング =====

/**
 * オークファンのHTMLを取得
 */
function fetchAucfanHtml_(url) {
  try {
    var options = getCommonHttpOptions_();
    var response = UrlFetchApp.fetch(url, options);

    validateHttpResponse_(response, "オークファン");
    return getResponseTextWithBestCharset_(response);
  } catch (e) {
    throw new Error("オークファン URLからのHTMLフェッチエラー: " + e.message);
  }
}

/**
 * オークファンのHTMLをパースして商品情報を抽出
 */
function parseAucfanFromHtml_(html) {
  const items = [];
  const base = "https://pro.aucfan.com";
  const re = /<li class="item">\s*<ul>([\s\S]*?)<\/ul>\s*<\/li>/gi;
  let m;

  while ((m = re.exec(html)) !== null) {
    const block = m[1];

    // 詳細URL
    const detailPath = firstMatch_(
      block,
      /<li class="title">[\s\S]*?<a\s+href="([^"]+)"/i
    );
    const detailUrl = detailPath
      ? detailPath.startsWith("http")
        ? detailPath
        : base + detailPath
      : "";

    // 画像URL
    const imageUrl = firstMatch_(
      block,
      /<img[^>]+class="thumbnail"[^>]+src="([^"]+)"/i
    );

    // タイトル抽出（複数パターンに対応）
    let title = "";
    const titleHtmlItemName = firstMatch_(
      block,
      /<li class="itemName">[\s\S]*?<a[^>]*>([\s\S]*?)<\/a>/i
    );
    if (titleHtmlItemName) {
      title = stripTags_(htmlDecode_(titleHtmlItemName))
        .replace(/\s+/g, " ")
        .trim();
    }

    // タイトルが取得できない場合の代替パターン
    if (!title) {
      const liItemNameAll = firstMatch_(
        block,
        /<li class="itemName">([\s\S]*?)<\/li>/i
      );
      if (liItemNameAll) {
        title = stripTags_(htmlDecode_(liItemNameAll))
          .replace(/\s+/g, " ")
          .trim();
      }
    }

    // 価格抽出（落札価格を優先）
    let price = "";

    // まず落札価格を探す（落札価格を優先）
    // パターン1: <span>落札</span>の後の価格
    let endPrice = firstMatch_(
      block,
      /<li class="price"[^>]*>\s*<span[^>]*>落札<\/span>\s*([^<]+)/i
    );

    // パターン2: 落札価格として明記されている場合
    if (!endPrice) {
      endPrice = firstMatch_(
        block,
        /落札価格[:：]?\s*([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i
      );
    }

    // パターン3: 終了価格として表示されている場合
    if (!endPrice) {
      endPrice = firstMatch_(
        block,
        /終了価格[:：]?\s*([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i
      );
    }

    if (endPrice) {
      // 落札価格が見つかった場合はそれを使用
      const normalizedEndPrice = normalizeNumber_(endPrice);
      if (normalizedEndPrice && normalizedEndPrice !== "0") {
        price = normalizedEndPrice;
      }
    }

    // 落札価格が見つからない場合は、既存の価格抽出ロジックを使用
    if (!price) {
      const priceCandidates = [];
      const pricePrimary = firstMatch_(
        block,
        /<li class="price"[^>]*>\s*([^<]+)/i
      );
      if (pricePrimary) priceCandidates.push(pricePrimary);

      const dataPrice = firstMatch_(block, /data-price="([0-9,]+)"/i);
      if (dataPrice) priceCandidates.push(dataPrice);

      const priceValue = firstMatch_(
        block,
        /class="[^"]*price__value[^"]*"[^>]*>\s*([^<]+)/i
      );
      if (priceValue) priceCandidates.push(priceValue);

      const yenInline = firstMatch_(block, /([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i);
      if (yenInline) priceCandidates.push(yenInline);

      const yenMark = firstMatch_(block, /¥\s*([0-9]{1,3}(?:,[0-9]{3})*)/i);
      if (yenMark) priceCandidates.push(yenMark);

      const nums = priceCandidates
        .map(function (p) {
          return normalizeNumber_(p);
        })
        .filter(function (n) {
          return n && n !== "0";
        })
        .map(function (n) {
          return parseInt(n, 10);
        })
        .filter(function (n) {
          return !isNaN(n) && n >= 100;
        });

      if (nums.length > 0) {
        // 明らかに送料や手数料と思われる金額を除外
        var filtered = nums.filter(function (n) {
          return (
            [
              3980, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000,
            ].indexOf(n) === -1
          );
        });
        var candidates = filtered.length > 0 ? filtered : nums;
        candidates.sort(function (a, b) {
          return a - b;
        });
        price = String(candidates[0]);
      }
    }

    // 終了日
    const endTxt = firstMatch_(block, /<li class="end">\s*([^<]+)/i);

    if (detailUrl || imageUrl || price || endTxt || title) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: endTxt || "",
          rank: "",
          price: price || "",
          shop: "",
          source: "オークファン",
          soldOut: false,
        })
      );
    }
  }

  return items;
}

/**
 * オークファンデータを取得（HTMLまたはURLから）
 */
function getAucfanDataFromSheetOrUrl_(sheet, urlRow, urlCol) {
  try {
    // まず67行目以降にHTMLがあるかチェック
    var htmlFromSheet = readHtmlFromRow_(sheet, 67);

    if (htmlFromSheet) {
      // HTMLソースがオークファンかどうかを判定
      var source = detectSource_(htmlFromSheet);
      if (source === "オークファン" || source === "aucfan") {
        console.log("67行目以降のHTMLからオークファンデータを取得します");
        var items = parseAucfanFromHtml_(htmlFromSheet);
        console.log("オークファンデータを取得しました:", items.length + "件");
        return items;
      } else {
        console.log("67行目以降のHTMLはオークファンではありません:", source);
      }
    }

    // HTMLがない場合はURLから取得を試行
    var url = sheet.getRange(urlRow || 110, urlCol || 2).getValue();
    if (url && url.toString().trim()) {
      var urlStr = url.toString().trim();
      if (urlStr.startsWith("http")) {
        console.log("URLからオークファンデータを取得します:", urlStr);
        var html = fetchAucfanHtml_(urlStr);
        var items = parseAucfanFromHtml_(html);
        console.log("オークファンデータを取得しました:", items.length + "件");
        return items;
      }
    }

    console.log("オークファンのHTMLもURLも見つかりません");
    return [];
  } catch (e) {
    console.warn("オークファンデータの取得に失敗:", e.message);
    return [];
  }
}

/**
 * オークファンHTML読み取りテスト関数
 */
function testAucfanHtmlFromSheet() {
  console.log("=== オークファンHTML読み取りテスト ===");

  try {
    var sheet = SpreadsheetApp.getActiveSheet();

    // 67行目以降のHTMLを読み取り
    var htmlFromSheet = readHtmlFromRow_(sheet, 67);

    if (!htmlFromSheet) {
      console.log("❌ 67行目以降にHTMLが見つかりません");
      console.log("B67行目以降にオークファンのHTMLを貼り付けてください");
      return;
    }

    console.log("✅ HTMLを読み取りました:", htmlFromSheet.length + "文字");

    // ソース判定
    var source = detectSource_(htmlFromSheet);
    console.log("判定されたソース:", source);

    if (source !== "オークファン") {
      console.log("⚠️ オークファンのHTMLではない可能性があります");
    }

    // パース実行
    var items = parseAucfanFromHtml_(htmlFromSheet);
    console.log("✅ パース完了:", items.length + "件の商品を取得");

    // 最初の3件を表示
    for (var i = 0; i < Math.min(3, items.length); i++) {
      console.log("\n商品" + (i + 1) + ":");
      console.log(
        "  タイトル:",
        items[i].title ? items[i].title.substring(0, 50) + "..." : "なし"
      );
      console.log("  価格:", items[i].price || "なし");
      console.log("  日付:", items[i].date || "なし");
    }

    if (items.length === 0) {
      console.log("❌ 商品データが取得できませんでした");
      console.log("HTMLの形式を確認してください");
    } else {
      console.log("\n🎉 テスト成功！");
    }
  } catch (e) {
    console.error("❌ テスト中にエラーが発生:", e.message);
    console.error("スタックトレース:", e.stack);
  }
}
