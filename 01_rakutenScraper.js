/**
 * 楽天市場専用スクレイパー
 * 楽天市場からの商品情報取得処理
 */

// ===== 楽天スクレイピング =====

/**
 * 現在のキーワードを取得（G7セルから）
 */
function getCurrentKeyword_() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var keyword = sheet.getRange("G7").getValue();
    return keyword ? String(keyword).trim() : "";
  } catch (e) {
    console.log("キーワード取得エラー:", e.message);
    return "";
  }
}

/**
 * 楽天の中古品状態テキストをランクに変換
 */
function convertConditionToRank_(conditionText) {
  if (!conditionText) return "";

  var condition = conditionText.toLowerCase().trim();

  // 状態テキストとランクのマッピング
  var conditionMapping = {
    未使用品: "S",
    ほぼ新品: "SA",
    非常に良い: "A",
    良い: "AB",
    可: "B",
  };

  // 完全一致を先にチェック
  for (var key in conditionMapping) {
    if (condition === key.toLowerCase()) {
      return conditionMapping[key];
    }
  }

  // 部分一致をチェック
  for (var key in conditionMapping) {
    if (condition.indexOf(key.toLowerCase()) !== -1) {
      return conditionMapping[key];
    }
  }

  console.log("楽天ランク変換失敗: 未知の状態テキスト:", conditionText);
  return ""; // マッピングにない場合は空白
}

/**
 * URLから検索キーワードを抽出
 */
function extractKeywordFromUrl_(url) {
  try {
    // URLから商品名部分を抽出
    var mallMatch = url.match(/\/search\/mall\/([^\/\?]+)/);
    if (mallMatch) {
      return decodeURIComponent(mallMatch[1]);
    }

    // qパラメータから抽出
    var qMatch = url.match(/[?&]q=([^&]+)/);
    if (qMatch) {
      return decodeURIComponent(qMatch[1]);
    }

    return "";
  } catch (e) {
    console.log("URL解析エラー:", e.message);
    return "";
  }
}

/**
 * 動的な推奨URL生成
 */
function generateRecommendedRakutenUrl_(currentUrl) {
  // 現在のキーワードを取得（G7セルから）
  var keyword = getCurrentKeyword_();

  // G7が空の場合はURLから抽出を試行
  if (!keyword) {
    keyword = extractKeywordFromUrl_(currentUrl);
  }

  // それでも取得できない場合はデフォルト
  if (!keyword) {
    keyword = "商品名";
  }

  // 推奨URLを生成（元の仕様に準拠）
  var encodedKeyword = encodeURIComponent(keyword);
  var recommendedUrl =
    "https://search.rakuten.co.jp/search/mall/" + encodedKeyword + "/?f=0&s=4";

  console.log("推奨URL生成:", {
    元キーワード: keyword,
    エンコード後: encodedKeyword,
    推奨URL: recommendedUrl,
  });

  return recommendedUrl;
}

/**
 * 楽天のURLが正しい条件（売り切れを含む、新しい順）になっているかチェック
 */
function validateRakutenSearchUrl_(url) {
  console.log("楽天URL検証開始:", url);

  var warnings = [];

  // 売り切れを含む条件をチェック（f=0 または f パラメータなし）
  var includesSoldOut = false;
  if (url.indexOf("f=0") !== -1 || url.indexOf("f=") === -1) {
    includesSoldOut = true;
  }

  if (!includesSoldOut) {
    warnings.push("売り切れを含む設定になっていません（f=0が必要）");
  }

  // 新しい順（新着順）をチェック（s=4）
  var isNewOrder = url.indexOf("s=4") !== -1;
  if (!isNewOrder) {
    warnings.push("新しい順（新着順）になっていません（s=4が必要）");
  }

  // 中古品に関する設定はオプション（元の仕様では指定なし）
  // var includesUsed = url.indexOf("used=1") !== -1;
  // if (!includesUsed) {
  //   warnings.push("中古品を含む設定になっていません（used=1が必要）");
  // }

  // 警告がある場合はアラートを表示
  if (warnings.length > 0) {
    var alertMessage = "⚠️ 楽天検索URLの設定に問題があります:\n\n";
    for (var i = 0; i < warnings.length; i++) {
      alertMessage += "• " + warnings[i] + "\n";
    }
    alertMessage += "\n現在のURL: " + url;

    // 動的に推奨URLを生成
    var recommendedUrl = generateRecommendedRakutenUrl_(url);
    alertMessage += "\n\n推奨URL: " + recommendedUrl;
    alertMessage += "";

    console.log("楽天URL警告:", alertMessage);

    // SpreadsheetApp.getUi()でアラート表示
    try {
      SpreadsheetApp.getUi().alert(
        "楽天検索URL設定の警告",
        alertMessage,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e) {
      console.log("アラート表示エラー（テスト環境の可能性）:", e.message);
    }
  } else {
    console.log("楽天URL検証OK: 正しい設定です");
  }

  return warnings.length === 0;
}

/**
 * 楽天のHTMLを取得
 */
function fetchRakutenHtml_(url) {
  try {
    // URL検証を実行（警告は出すが処理は続行）
    validateRakutenSearchUrl_(url);

    var options = getCommonHttpOptions_();
    var response = UrlFetchApp.fetch(url, options);

    validateHttpResponse_(response, "楽天");
    return getResponseTextWithBestCharset_(response);
  } catch (e) {
    throw new Error("楽天URLからのHTMLフェッチエラー: " + e.message);
  }
}

/**
 * 楽天の商品ブロックから商品情報を抽出
 */
function extractRakutenItemFromBlock_(block, detailUrl, index) {
  // 画像URL抽出
  var imageUrl = "";
  var imgPatterns = [
    // 基本的なsrc属性
    /<img[^>]+src="([^"]+)"/i,
    // 遅延読み込み用のdata-src
    /<img[^>]+data-src="([^"]+)"/i,
    // その他のdata属性
    /<img[^>]+data-original="([^"]+)"/i,
    /<img[^>]+data-lazy="([^"]+)"/i,
    /<img[^>]+data-image="([^"]+)"/i,
    // 楽天特有のパターン
    /<img[^>]+data-rakuten-src="([^"]+)"/i,
    // より広範囲なimg要素の検索
    /<img[^>]*src\s*=\s*["']([^"']+)["'][^>]*>/i,
    // 楽天商品画像の典型的なURL構造
    /https?:\/\/[^"'\s]*\.rakuten\.co\.jp[^"'\s]*\.(jpg|jpeg|png|gif|webp)/i,
    // 楽天CDNの画像URL
    /https?:\/\/[^"'\s]*rakuten[^"'\s]*\.(jpg|jpeg|png|gif|webp)/i,
    // 楽天の新しい画像パターン
    /<img[^>]+data-lazy-src="([^"]+)"/i,
    /<img[^>]+data-img="([^"]+)"/i,
    // より柔軟な画像URL検索
    /["']([^"']*(?:thumbnail|image|img)[^"']*\.(?:jpg|jpeg|png|gif|webp))[^"']*/i,
    // 楽天の画像サーバーパターン
    /["'](https?:\/\/(?:thumbnail|image|img)\.rakuten\.co\.jp[^"']*\.(?:jpg|jpeg|png|gif|webp))[^"']*/i,
    // より広範囲な楽天画像URL
    /["'](https?:\/\/[^"']*rakuten[^"']*\/[^"']*\.(?:jpg|jpeg|png|gif|webp))[^"']*/i,
  ];

  for (var i = 0; i < imgPatterns.length; i++) {
    var pattern = imgPatterns[i];
    var imgMatch = block.match(pattern);
    if (imgMatch) {
      var extractedUrl = htmlDecode_(imgMatch[1]);
      console.log("楽天画像URL候補 (パターン" + (i + 1) + "):", extractedUrl);

      // URLの妥当性チェック
      if (
        extractedUrl &&
        extractedUrl.startsWith("http") &&
        !extractedUrl.includes("noimage") &&
        !extractedUrl.includes("blank") &&
        !extractedUrl.includes("placeholder") &&
        !extractedUrl.includes("loading") &&
        (extractedUrl.includes(".jpg") ||
          extractedUrl.includes(".jpeg") ||
          extractedUrl.includes(".png") ||
          extractedUrl.includes(".gif") ||
          extractedUrl.includes(".webp"))
      ) {
        imageUrl = extractedUrl;
        console.log("楽天画像URL抽出成功 (パターン" + (i + 1) + "):", imageUrl);
        break;
      } else {
        console.log("楽天画像URL無効 (パターン" + (i + 1) + "):", extractedUrl);
      }
    }
  }

  // デバッグ用ログ
  if (!imageUrl) {
    console.log(
      "楽天画像URL抽出失敗 - ブロック内容の一部:",
      block.substring(0, 500)
    );
  }

  // 商品名抽出
  var title = "";
  var titlePatterns = [
    // 楽天の新しいUI構造に対応 - より正確なパターン
    /<a[^>]+class="[^"]*title-link--[^"]*"[^>]*>([^<]+)<\/a>/i,
    // title属性から取得（最も確実）
    /<a[^>]+title="([^"]+)"[^>]+class="[^"]*title-link/i,
    // data-link="item"を含むパターン
    /<a[^>]+data-link="item"[^>]+>([^<]+)<\/a>/i,
    // より広範囲なtitle-linkパターン
    /<a[^>]*class="[^"]*title-link[^"]*"[^>]*>([\s\S]*?)<\/a>/i,
    // 従来のパターン
    /<a[^>]+href="https?:\/\/item\.rakuten\.co\.jp[^>]*>([^<]+)<\/a>/i,
  ];

  for (var pattern of titlePatterns) {
    var titleMatch = block.match(pattern);
    if (titleMatch) {
      title = stripTags_(htmlDecode_(titleMatch[1]))
        .replace(/\s+/g, " ")
        .replace(/\n/g, " ")
        .trim();
      if (title) {
        console.log("楽天タイトル抽出成功:", title.substring(0, 50) + "...");
        break;
      }
    }
  }

  if (!title) {
    console.log(
      "楽天タイトル抽出失敗 - ブロック内容の一部:",
      block.substring(0, 300)
    );
  }

  // 価格抽出
  var price = "";

  // まず、data-track-price属性から価格を取得を試みる
  var trackPriceMatch = block.match(/data-track-price="([0-9]+)"/i);
  if (trackPriceMatch) {
    price = trackPriceMatch[1];
    console.log("楽天価格抽出成功 (data-track-price):", price);
  }

  // data-track-priceが見つからない場合は、他のパターンを試す
  if (!price) {
    var pricePatterns = [
      // 新しい楽天の価格構造に対応
      /<div[^>]+class="[^"]*price--[^"]*"[^>]*>([0-9]{1,3}(?:,[0-9]{3})*)<span/i,
      // より一般的な価格divパターン
      /<div[^>]+class="[^"]*price[^"]*"[^>]*>([0-9]{1,3}(?:,[0-9]{3})*)/i,
      // 価格spanパターン
      /<span[^>]+class="[^"]*price[^"]*"[^>]*>([0-9]{1,3}(?:,[0-9]{3})*)/i,
      // 円を含むパターン
      /([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i,
      // ¥記号を含むパターン
      /¥\s*([0-9]{1,3}(?:,[0-9]{3})*)/i,
    ];

    for (var i = 0; i < pricePatterns.length; i++) {
      var pattern = pricePatterns[i];
      var priceMatch = block.match(pattern);
      if (priceMatch) {
        var normalizedPrice = normalizeNumber_(priceMatch[1]);
        if (normalizedPrice && parseInt(normalizedPrice) >= 100) {
          price = normalizedPrice;
          console.log("楽天価格抽出成功 (パターン" + (i + 1) + "):", price);
          break;
        }
      }
    }
  }

  if (!price) {
    console.log(
      "楽天価格抽出失敗 - ブロック内容の一部:",
      block.substring(0, 300)
    );
  }

  // ランク抽出（中古品の状態から判定）
  var rank = "";
  var rankPatterns = [
    // 中古品の状態を示すテキストパターン
    /<div[^>]+class="[^"]*used-condition[^"]*"[^>]*>[\s\S]*?<div[^>]+class="[^"]*text-display[^"]*"[^>]*>([^<]+)<\/div>/i,
    // より広範囲なパターン
    /中古品-([^<]+)/i,
    // 状態を示すテキストの直接検索
    /(?:中古品-|状態：?)([^<\s]+)/i,
  ];

  for (var i = 0; i < rankPatterns.length; i++) {
    var pattern = rankPatterns[i];
    var rankMatch = block.match(pattern);
    if (rankMatch) {
      var conditionText = rankMatch[1].trim();
      console.log("楽天状態テキスト抽出:", conditionText);

      // 状態テキストをランクに変換
      rank = convertConditionToRank_(conditionText);
      if (rank) {
        console.log("楽天ランク変換成功:", conditionText + " → " + rank);
        break;
      }
    }
  }

  if (!rank) {
    console.log(
      "楽天ランク抽出失敗 - ブロック内容の一部:",
      block.substring(0, 300)
    );
  }

  // 売り切れ情報抽出
  var soldOut = false;
  var soldOutPatterns = [
    // より柔軟なパターン - classの順序を問わない
    /<span[^>]+class="[^"]*status--22u55[^"]*"[^>]*>売り切れ<\/span>/i,
    // content statusを含むdivの中の売り切れ
    /<div[^>]+class="[^"]*content[^"]*status[^"]*"[^>]*>[\s\S]*?売り切れ[\s\S]*?<\/div>/i,
    // さらに広いパターン
    /class="status--22u55"[^>]*>売り切れ</i,
    // 最も基本的なパターン
    />売り切れ</,
    // 在庫なしパターン
    />[\s\S]*?在庫なし[\s\S]*?</i,
    /在庫なし/i,
    // 品切れパターン
    />[\s\S]*?品切れ[\s\S]*?</i,
    /品切れ/i,
    // SOLD OUTパターン
    />[\s\S]*?SOLD[\s\S]*?OUT[\s\S]*?</i,
    /SOLD[\s\S]*?OUT/i,
  ];

  // 売り切れ検出デバッグ
  var blockPreview = block.substring(0, 500);
  var hasStatusDiv = block.indexOf('class="content status"') !== -1;
  var hasSoldOutText = block.indexOf("売り切れ") !== -1;

  console.log("楽天売り切れ検出開始 - 商品" + (index + 1) + ":", {
    hasStatusDiv: hasStatusDiv,
    hasSoldOutText: hasSoldOutText,
    ブロック長: block.length,
  });

  for (var i = 0; i < soldOutPatterns.length; i++) {
    var pattern = soldOutPatterns[i];
    if (block.match(pattern)) {
      soldOut = true;
      console.log(
        "楽天売り切れ検出成功! パターン" + (i + 1) + ":",
        pattern.toString()
      );
      break;
    }
  }

  if (!soldOut && hasSoldOutText) {
    console.log(
      "警告: 売り切れテキストは存在するが、パターンにマッチしませんでした"
    );
    console.log("ブロック内容の一部:", blockPreview);
  }

  if (!soldOut) {
    console.log("楽天売り切れ未検出 - 在庫ありと判定");
  }

  // デバッグ用：最初の商品を強制的に売り切れにする（コメントアウト）
  // if (index === 0 && title) {
  //   soldOut = true;
  //   console.log("★★★ デバッグ：最初の商品を強制的に売り切れに設定 ★★★");
  // }

  var itemData = createItemData_({
    title: title,
    detailUrl: detailUrl,
    imageUrl: imageUrl,
    date: "",
    rank: rank,
    price: price,
    shop: "",
    source: "楽天",
    soldOut: soldOut,
  });

  console.log("楽天商品データ作成完了:", {
    title: title ? title.substring(0, 30) + "..." : "なし",
    rank: rank,
    soldOut: soldOut,
    soldOutプロパティ: itemData.soldOut,
  });

  // createItemData_の戻り値を詳細確認
  console.log("★★★ createItemData_戻り値詳細確認 ★★★");
  console.log("itemData:", itemData);
  console.log("itemData.soldOut:", itemData.soldOut);
  console.log("itemData全プロパティ:", Object.keys(itemData));
  console.log("itemDataのJSON:", JSON.stringify(itemData));

  return itemData;
}

/**
 * 楽天のHTMLをパースして商品情報を抽出
 */
function parseRakutenFromHtml_(html) {
  const items = [];
  const H = String(html);

  // 商品ブロックを抽出するパターン - より正確なパターンを先頭に配置
  const blockPatterns = [
    // 楽天の新しいUI（dui-card searchresultitem）
    /<div[^>]+class="dui-card searchresultitem[^"]*"[^>]*>[\s\S]*?(?=<div[^>]+class="dui-card searchresultitem|$)/gi,
  ];

  let productBlocks = [];

  // 各パターンを試して商品ブロックを抽出
  for (let i = 0; i < blockPatterns.length; i++) {
    const pattern = blockPatterns[i];
    const matches = H.match(pattern);
    if (matches && matches.length > 0) {
      productBlocks = matches;
      console.log(
        "楽天商品ブロック検出: パターン" +
          (i + 1) +
          "で" +
          matches.length +
          "件"
      );
      break;
    }
  }

  // ブロックベースで抽出できた場合
  if (productBlocks.length > 0) {
    productBlocks.forEach(function (block, index) {
      const detailUrlMatch = block.match(
        /<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/i
      );
      if (!detailUrlMatch) return;

      const detailUrl = htmlDecode_(detailUrlMatch[1]);
      items.push(extractRakutenItemFromBlock_(block, detailUrl, index));
    });
  } else {
    // リンクベース方式にフォールバック（元の実装）
    console.log("ブロック検出失敗。リンクベース方式にフォールバック");

    // 楽天商品リンクを全て見つける
    const linkPattern = /https?:\/\/item\.rakuten\.co\.jp\/[^"]+/gi;
    const links = [];
    let match;

    while ((match = linkPattern.exec(H)) !== null) {
      links.push({
        url: match[0],
        startIndex: match.index,
        endIndex: match.index + match[0].length,
      });
    }

    // 各リンクを中心に前後のHTMLを取得
    links.forEach(function (link, index) {
      const searchStart = Math.max(0, link.startIndex - 2500);
      const searchEnd = Math.min(H.length, link.endIndex + 2500);
      const block = H.substring(searchStart, searchEnd);

      const detailUrl = htmlDecode_(link.url);
      items.push(extractRakutenItemFromBlock_(block, detailUrl, index));
    });
  }

  // 最終的な商品データの確認
  console.log("=== 楽天パーサー最終結果 ===");
  console.log("商品数:", items.length);
  items.slice(0, 3).forEach(function (item, index) {
    console.log("最終商品" + (index + 1) + ":", {
      title: item.title ? item.title.substring(0, 30) + "..." : "なし",
      rank: item.rank,
      soldOut: item.soldOut,
      soldOutプロパティ存在: item.hasOwnProperty("soldOut"),
      全プロパティ: Object.keys(item),
    });
  });

  return items;
}

/**
 * 楽天の詳細ページから商品情報を抽出
 */
function parseRakutenDetailFromHtml_(html, detailUrl) {
  // 詳細ページの解析ロジック（必要に応じて実装）
  return createItemData_({
    title: "",
    detailUrl: detailUrl,
    imageUrl: "",
    date: "",
    rank: "",
    price: "",
    shop: "",
    source: "楽天",
    soldOut: false,
  });
}
