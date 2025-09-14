/**
 * Star Buyers Auction専用スクレイパー
 * SBAからの商品情報取得処理（ログイン処理含む）
 */

// ===== Star Buyers Auction スクレイピング =====

/**
 * Star Buyers AuctionのHTMLを取得（ログイン処理含む）
 */
function fetchStarBuyersHtml_(url) {
  try {
    // ランダムな待機時間（1-3秒）
    Utilities.sleep(1000 + Math.floor(Math.random() * 2000));

    // ログイン情報
    var loginEmail = "inui.hur@gmail.com";
    var loginPassword = "hur22721";
    var loginUrl = "https://www.starbuyers-global-auction.com/login";

    // ログインページにアクセス
    var loginPageOptions = {
      method: "get",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
        Accept:
          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
      },
      muteHttpExceptions: true,
      followRedirects: false,
    };

    var loginPageResponse = UrlFetchApp.fetch(loginUrl, loginPageOptions);
    var loginPageHtml = loginPageResponse.getContentText("UTF-8");
    var cookies = loginPageResponse.getAllHeaders()["Set-Cookie"];

    // CSRFトークンを抽出
    var csrfToken = "";
    var csrfMatch = loginPageHtml.match(
      /<input[^>]+name="_token"[^>]+value="([^"]+)"/
    );
    if (!csrfMatch) {
      csrfMatch = loginPageHtml.match(
        /<meta name="csrf-token" content="([^"]+)"/
      );
    }
    if (csrfMatch) {
      csrfToken = csrfMatch[1];
    }

    // クッキーを処理
    var cookieMap = {};
    if (cookies) {
      if (Array.isArray(cookies)) {
        cookies.forEach(function (cookie) {
          var parts = cookie.split(";")[0].split("=");
          if (parts.length >= 2) {
            cookieMap[parts[0]] = parts.slice(1).join("=");
          }
        });
      } else {
        var parts = cookies.split(";")[0].split("=");
        if (parts.length >= 2) {
          cookieMap[parts[0]] = parts.slice(1).join("=");
        }
      }
    }

    // ログイン実行
    var loginPayload = {
      email: loginEmail,
      password: loginPassword,
      _token: csrfToken,
    };

    var cookieString = Object.keys(cookieMap)
      .map(function (key) {
        return key + "=" + cookieMap[key];
      })
      .join("; ");

    var payloadString = Object.keys(loginPayload)
      .map(function (key) {
        return (
          encodeURIComponent(key) + "=" + encodeURIComponent(loginPayload[key])
        );
      })
      .join("&");

    var loginOptions = {
      method: "post",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
        "Content-Type": "application/x-www-form-urlencoded",
        Cookie: cookieString,
        Referer: loginUrl,
        Origin: "https://www.starbuyers-global-auction.com",
      },
      payload: payloadString,
      muteHttpExceptions: true,
      followRedirects: false,
    };

    var loginResponse = UrlFetchApp.fetch(loginUrl, loginOptions);

    // ログイン後のクッキーを取得
    var loginCookies = loginResponse.getAllHeaders()["Set-Cookie"];
    if (loginCookies) {
      if (Array.isArray(loginCookies)) {
        loginCookies.forEach(function (cookie) {
          var parts = cookie.split(";")[0].split("=");
          if (parts.length >= 2) {
            cookieMap[parts[0]] = parts.slice(1).join("=");
          }
        });
      }
    }

    // 実際のURLにアクセス
    var finalCookieString = Object.keys(cookieMap)
      .map(function (key) {
        return key + "=" + cookieMap[key];
      })
      .join("; ");

    var finalOptions = {
      method: "get",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
        Cookie: finalCookieString,
        Referer: "https://www.starbuyers-global-auction.com/",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        Accept:
          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
      },
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(url, finalOptions);

    if (response.getResponseCode() !== 200) {
      throw new Error(
        "SBAからのHTMLの取得に失敗しました。ステータスコード: " +
          response.getResponseCode()
      );
    }

    return response.getContentText("UTF-8");
  } catch (e) {
    throw new Error("SBA URLからのHTMLフェッチエラー: " + e.message);
  }
}

/**
 * Star Buyers AuctionのHTMLをパースして商品情報を抽出
 */
function parseStarBuyersFromHtml_(html) {
  var items = [];
  var H = String(html);

  console.log("SBAパーサー - HTMLの長さ:", H.length);
  console.log(
    "SBAパーサー - HTMLに'p-item-list__body'が含まれているか:",
    H.indexOf("p-item-list__body") > -1
  );
  console.log(
    "SBAパーサー - HTMLに'market_price'が含まれているか:",
    H.indexOf("market_price") > -1
  );

  // 各アイテムの開始インデックスを列挙
  var starts = [];
  var re = /<div[^>]*class="p-item-list__body"[^>]*>/g;
  var m;
  while ((m = re.exec(H)) !== null) {
    starts.push(m.index);
  }

  console.log("SBAパーサー - p-item-list__bodyが見つかった数:", starts.length);

  // フォールバック: 別のセレクターパターンを試す
  if (starts.length === 0) {
    console.log("SBAパーサー - フォールバック: 別のセレクターを試します");

    // より具体的なセレクターを試す
    var alternativeSelectors = [
      /<div[^>]*class="p-item-list__body">/g, // スペース要件を緩和
      /<div[^>]*p-item-list__body[^>]*>/g, // class属性の位置を問わない
    ];

    for (
      var selectorIndex = 0;
      selectorIndex < alternativeSelectors.length;
      selectorIndex++
    ) {
      var altRe = alternativeSelectors[selectorIndex];
      altRe.lastIndex = 0; // reset regex
      starts = []; // リセット
      while ((m = altRe.exec(H)) !== null) {
        starts.push(m.index);
      }
      if (starts.length > 0) {
        console.log(
          "SBAパーサー - 代替セレクター" +
            selectorIndex +
            "で" +
            starts.length +
            "件見つかりました"
        );
        break;
      }
    }
  }

  if (starts.length === 0) {
    console.log("SBAパーサー - HTMLサンプル（最初の1000文字）:");
    console.log(H.substring(0, 1000));

    // より詳細なデバッグ情報
    console.log("SBAパーサー - HTMLに含まれる主要なクラス名:");
    var classMatches = H.match(/class="[^"]*"/g);
    if (classMatches) {
      var uniqueClasses = {};
      for (var i = 0; i < Math.min(classMatches.length, 20); i++) {
        var className = classMatches[i];
        if (
          className.indexOf("item") > -1 ||
          className.indexOf("list") > -1 ||
          className.indexOf("body") > -1
        ) {
          uniqueClasses[className] = true;
        }
      }
      console.log("関連するクラス名:", Object.keys(uniqueClasses).join(", "));
    }

    // 商品データらしきパターンを探す
    var productPatterns = [
      /セリーヌ[^<]*ヴィクトワール/g,
      /market_price\/\d+/g,
      /\d+,\d+yen/g,
      /<a[^>]*href="[^"]*market_price[^"]*"[^>]*>/g,
    ];

    for (var p = 0; p < productPatterns.length; p++) {
      var matches = H.match(productPatterns[p]);
      if (matches) {
        console.log("パターン" + p + "のマッチ数:", matches.length);
        if (matches.length > 0) {
          console.log("最初のマッチ:", matches[0]);
        }
      }
    }

    // フォールバック: 商品データが見つかった場合、汎用的な抽出を試行
    var productKeyword = /セリーヌ[^<]*ヴィクトワール/;
    var productMatch = H.match(productKeyword);
    if (productMatch) {
      console.log(
        "SBAパーサー - 商品キーワードが見つかりました、汎用抽出を試行"
      );

      // まず、sba.htmlファイルの内容かどうかをチェック
      if (H.indexOf('class="p-item-list__body"') > -1) {
        console.log("SBAパーサー - sba.htmlファイルを検出、専用パーサーを使用");
        return parseSbaHtmlFile_(H);
      }

      return extractProductsFromGenericHtml_(H);
    }

    return items;
  }

  // 開始インデックスごとに商品情報を抽出
  for (var i = 0; i < starts.length; i++) {
    var s = starts[i];
    var e = i + 1 < starts.length ? starts[i + 1] : H.length;
    var block = H.slice(s, e);

    console.log(
      "SBAパーサー - アイテム" + (i + 1) + "/" + starts.length + "を処理中"
    );

    // 詳細URL
    var detailUrl = "";
    var mUrl = block.match(
      /<a[^>]+href="(https:\/\/www\.starbuyers-global-auction\.com\/market_price\/[^"]+)"/i
    );
    if (mUrl) detailUrl = htmlDecode_(mUrl[1]);

    // 画像URL
    var imageUrl = "";
    var mImg = block.match(/<img[^>]+src="([^"]+)"[^>]*>/i);
    if (mImg) imageUrl = htmlDecode_(mImg[1]);

    // 商品名（p-text-linkから抽出）
    var title = "";
    var mTitle = block.match(
      /<a[^>]+class="p-text-link"[^>]*href="[^"]*market_price[^"]*"[^>]*>([\s\S]*?)<\/a>/i
    );
    if (mTitle) {
      title = stripTags_(htmlDecode_(mTitle[1])).replace(/\s+/g, " ").trim();
    }

    // 開催日
    var date = "";
    var mDate = block.match(
      /data-head="開催日"[\s\S]*?<strong>([^<]+)<\/strong>/i
    );
    if (mDate) date = mDate[1].trim();

    // ランク
    var rank = "";
    var mRank = block.match(/data-rank="([^"]+)"/i);
    if (mRank) {
      var rawRank = mRank[1].trim();
      rank = normalizeRank_(rawRank);
      // ランクのデバッグ
      if (rawRank.toUpperCase() === "N" || rank === "N") {
        console.log(
          "SBA - ランクN検出: raw='" + rawRank + "' → normalized='" + rank + "'"
        );
      }
      // SAランクのデバッグ
      if (rawRank === "ＳＡ" || rawRank.toUpperCase() === "SA") {
        console.log(
          "SBA - SAランク検出: raw='" +
            rawRank +
            "' → normalized='" +
            rank +
            "'"
        );
      }
    }

    // 価格（"96,000yen" 等）
    var price = "";
    var mPrice = block.match(
      /data-head="落札額"[\s\S]*?<strong>([^<]+)<\/strong>/i
    );
    if (mPrice) {
      var raw = mPrice[1].trim();
      var num = normalizeNumber_(raw.replace(/yen/i, ""));
      // 数値化できた場合は数値、できなかった場合は元のテキストを保持
      price = num || raw;
      console.log("SBA価格抽出:", raw, "→", price);
    }

    console.log("SBAパーサー - アイテム" + (i + 1) + "の抽出結果:", {
      title: title ? "あり" : "なし",
      detailUrl: detailUrl ? "あり" : "なし",
      imageUrl: imageUrl ? "あり" : "なし",
      date: date ? "あり" : "なし",
      rank: rank ? "あり" : "なし",
      price: price ? "あり" : "なし",
    });

    if (detailUrl || imageUrl || date || rank || price || title) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: date,
          rank: rank,
          price: price,
          shop: "",
          source: "SBA",
          soldOut: false,
        })
      );
      console.log("SBAパーサー - アイテム" + (i + 1) + "を追加しました");
    } else {
      console.log(
        "SBAパーサー - アイテム" +
          (i + 1) +
          "は有効なデータがないためスキップしました"
      );
    }
  }

  console.log(
    "SBAパーサー - 最終結果:",
    items.length + "件のアイテムを抽出しました"
  );
  return items;
}

/**
 * sba.htmlファイル専用のパーサー
 */
function parseSbaHtmlFile_(html) {
  var items = [];
  var H = String(html);

  console.log("SBAパーサー - sba.htmlファイル専用パーサーを開始");

  // sba.htmlの正確な構造に基づいて商品ブロックを抽出
  var itemBlocks = [];
  var bodyRegex =
    /<div class="p-item-list__body">([\s\S]*?)(?=<div class="p-item-list__body">|$)/g;
  var match;

  while ((match = bodyRegex.exec(H)) !== null) {
    itemBlocks.push(match[1]);
  }

  console.log("SBAパーサー - 商品ブロック数:", itemBlocks.length);

  for (var i = 0; i < itemBlocks.length; i++) {
    var block = itemBlocks[i];

    // 商品名を抽出
    var titleMatch = block.match(
      /<a class="p-text-link" href="[^"]*"[^>]*>([^<]+)<\/a>/
    );
    var title = titleMatch ? titleMatch[1].trim() : "";

    // 詳細URLを抽出
    var urlMatch = block.match(
      /<a[^>]+href="(https:\/\/www\.starbuyers-global-auction\.com\/market_price\/[^"]+)"/
    );
    var detailUrl = urlMatch ? urlMatch[1] : "";

    // 画像URLを抽出
    var imageMatch = block.match(/<img[^>]+src="([^"]+)"/);
    var imageUrl = imageMatch ? imageMatch[1] : "";

    // 価格を抽出
    var priceMatch = block.match(
      /data-head="落札額"[\s\S]*?<strong>([^<]+)<\/strong>/
    );
    var price = "";
    if (priceMatch) {
      price = priceMatch[1].replace(/yen/i, "").trim();
    }

    // 開催日を抽出
    var dateMatch = block.match(
      /data-head="開催日"[\s\S]*?<strong>([^<]+)<\/strong>/
    );
    var date = dateMatch ? dateMatch[1].trim() : "";

    // ランクを抽出
    var rankMatch = block.match(/data-rank="([^"]+)"/);
    var rank = rankMatch ? rankMatch[1].trim() : "";

    console.log("SBAパーサー - アイテム" + (i + 1) + ":", {
      title: title || "なし",
      price: price || "なし",
      detailUrl: detailUrl ? "あり" : "なし",
      date: date || "なし",
      rank: rank || "なし",
    });

    if (title || detailUrl || price) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: date,
          rank: rank,
          price: price,
          shop: "",
          source: "SBA",
          soldOut: false,
        })
      );
    }
  }

  console.log(
    "SBAパーサー - sba.htmlファイル専用パーサー完了:",
    items.length + "件"
  );
  return items;
}

/**
 * 汎用的なHTML構造から商品情報を抽出（フォールバック用）
 */
function extractProductsFromGenericHtml_(html) {
  var items = [];
  var H = String(html);

  console.log("SBAパーサー - 汎用抽出を開始");

  // 商品名のパターンを探す（より広範囲なパターンを試行）
  var titlePatterns = [
    /セリーヌ[^<>]*ヴィクトワール[^<>]*/g,
    /セリーヌ[^<>]*トリオンフ[^<>]*/g,
    /セリーヌ[^<>]*レザー[^<>]*バッグ[^<>]*/g,
    /セリーヌ[^<>]*ショルダーバッグ[^<>]*/g,
    /セリーヌ[^<>]*チェーン[^<>]*/g,
  ];

  var foundTitles = [];
  for (var t = 0; t < titlePatterns.length; t++) {
    var titleMatches = H.match(titlePatterns[t]);
    if (titleMatches) {
      console.log(
        "商品名パターン" + t + "のマッチ:",
        titleMatches.length + "件"
      );
      for (var tm = 0; tm < Math.min(titleMatches.length, 2); tm++) {
        console.log("商品名例" + (tm + 1) + ":", titleMatches[tm]);
      }

      for (var m = 0; m < titleMatches.length; m++) {
        var cleanTitle = titleMatches[m].replace(/\s+/g, " ").trim();
        if (cleanTitle.length > 5 && foundTitles.indexOf(cleanTitle) === -1) {
          foundTitles.push(cleanTitle);
        }
      }
    }
  }

  // より汎用的な商品名パターンも試す
  if (foundTitles.length === 0) {
    console.log("SBAパーサー - 汎用的な商品名パターンを試行");
    var genericTitleMatches = H.match(
      /[^\s<>]{3,}[^<>]*(?:バッグ|ショルダー|レザー|チェーン)[^<>]*/g
    );
    if (genericTitleMatches) {
      for (var g = 0; g < Math.min(genericTitleMatches.length, 5); g++) {
        var genericTitle = genericTitleMatches[g].replace(/\s+/g, " ").trim();
        if (
          genericTitle.length > 8 &&
          foundTitles.indexOf(genericTitle) === -1
        ) {
          foundTitles.push(genericTitle);
          console.log("汎用商品名" + (g + 1) + ":", genericTitle);
        }
      }
    }
  }

  console.log("SBAパーサー - 見つかった商品名:", foundTitles.length + "件");

  // 価格パターンを探す（より広範囲なパターンを試行）
  var pricePatterns = [
    /\d{1,3}(?:,\d{3})*(?:yen|円)/g,
    /\d{1,3}(?:,\d{3})+/g, // カンマ区切りの数字
    /\d{6,}/g, // 6桁以上の数字（価格の可能性）
  ];

  var prices = [];
  for (var pp = 0; pp < pricePatterns.length; pp++) {
    var priceMatches = H.match(pricePatterns[pp]);
    if (priceMatches) {
      console.log(
        "価格パターン" + pp + "のマッチ:",
        priceMatches.length + "件"
      );
      for (var p = 0; p < Math.min(priceMatches.length, 3); p++) {
        console.log("価格例" + (p + 1) + ":", priceMatches[p]);
      }

      for (var p = 0; p < priceMatches.length; p++) {
        var price = priceMatches[p].replace(/yen|円/g, "").trim();
        if (price && prices.indexOf(price) === -1) {
          // 価格らしい数値をフィルタリング（10万円以上500万円以下）
          var numPrice = parseInt(price.replace(/,/g, ""));
          if (numPrice >= 100000 && numPrice <= 5000000) {
            prices.push(price);
          }
        }
      }
      if (prices.length > 0) break; // 価格が見つかったら他のパターンは試さない
    }
  }

  console.log("SBAパーサー - 見つかった価格:", prices.length + "件");

  // URLパターンを探す（より広範囲なパターンを試行）
  var urlPatterns = [
    /https:\/\/www\.starbuyers-global-auction\.com\/market_price\/\d+/g,
    /market_price\/\d+/g, // 相対パス
    /\/market_price\/\d+/g, // スラッシュ付き相対パス
  ];

  var urls = [];
  for (var up = 0; up < urlPatterns.length; up++) {
    var urlMatches = H.match(urlPatterns[up]);
    if (urlMatches) {
      console.log("URLパターン" + up + "のマッチ:", urlMatches.length + "件");
      for (var u = 0; u < Math.min(urlMatches.length, 3); u++) {
        console.log("URL例" + (u + 1) + ":", urlMatches[u]);
      }

      for (var u = 0; u < urlMatches.length; u++) {
        var url = urlMatches[u];
        // 相対パスの場合は絶対パスに変換
        if (!url.startsWith("http")) {
          url =
            "https://www.starbuyers-global-auction.com" +
            (url.startsWith("/") ? "" : "/") +
            url;
        }
        if (urls.indexOf(url) === -1) {
          urls.push(url);
        }
      }
      if (urls.length > 0) break; // URLが見つかったら他のパターンは試さない
    }
  }

  console.log("SBAパーサー - 見つかったURL:", urls.length + "件");

  // データを組み合わせて商品アイテムを作成
  var maxItems = Math.max(foundTitles.length, prices.length, urls.length);
  for (var i = 0; i < maxItems; i++) {
    var item = createItemData_({
      title: foundTitles[i] || "",
      detailUrl: urls[i] || "",
      imageUrl: "",
      date: "",
      rank: "",
      price: prices[i] || "",
      shop: "",
      source: "SBA",
      soldOut: false,
    });

    // 有効なデータがある場合のみ追加
    if (item.title || item.detailUrl || item.price) {
      items.push(item);
      console.log("SBAパーサー - アイテム" + (i + 1) + "を追加:", {
        title: item.title ? "あり" : "なし",
        detailUrl: item.detailUrl ? "あり" : "なし",
        price: item.price ? "あり" : "なし",
      });
    }
  }

  console.log("SBAパーサー - 汎用抽出完了:", items.length + "件");
  return items;
}

/**
 * Star Buyers Auctionの詳細ページから商品情報を抽出
 */
function parseSbaDetailFromHtml_(html, detailUrl) {
  let items = parseStarBuyersFromHtml_(html);
  return items.length > 0
    ? items[0]
    : createItemData_({
        title: "",
        detailUrl: detailUrl,
        imageUrl: "",
        date: "",
        rank: "",
        price: "",
        shop: "",
        source: "SBA",
        soldOut: false,
      });
}
