/**
 * エコリング専用スクレイパー
 * エコリングからの商品情報取得処理
 */

// ===== エコリング スクレイピング =====

/**
 * エコリングのHTMLを取得（ログイン必要）
 */
function fetchEcoringHtml_(url) {
  try {
    logInfo_("fetchEcoringHtml_", "エコリングHTML取得開始", "URL: " + url);
    // ランダムな待機時間（1-3秒）
    Utilities.sleep(1000 + Math.floor(Math.random() * 2000));

    // ログイン情報（ベタ打ち）
    var loginEmail = "info@genkaya.jp";
    var loginPassword = "ecoringenkaya";

    // ログインURL
    var loginUrl = "https://www.ecoauc.com/client/users/sign-in";

    // ステップ1: ログインページにアクセスしてクッキーを取得
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
    };

    var loginPageResponse = UrlFetchApp.fetch(loginUrl, loginPageOptions);
    var loginPageHtml = loginPageResponse.getContentText("UTF-8");
    var cookies = loginPageResponse.getAllHeaders()["Set-Cookie"];

    // CSRFトークンを抽出（エコリングは_csrfTokenを使用）
    var csrfToken = "";
    var csrfMatch = loginPageHtml.match(
      /<input[^>]+name="_csrfToken"[^>]+value="([^"]+)"/
    );
    if (!csrfMatch) {
      // 属性の順序が異なる場合
      csrfMatch = loginPageHtml.match(
        /<input[^>]+value="([^"]+)"[^>]+name="_csrfToken"/
      );
    }
    if (csrfMatch) {
      csrfToken = csrfMatch[1];
    }

    console.log("エコリング CSRFトークン取得:", csrfToken ? "成功" : "失敗");
    if (!csrfToken) {
      logWarning_(
        "fetchEcoringHtml_",
        "CSRFトークンが取得できませんでした",
        ""
      );
    }

    // セッションクッキーを収集
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

    // ステップ2: ログイン実行（エコリングのフォームに合わせる）
    var loginPayload = {
      _method: "POST",
      _csrfToken: csrfToken,
      email_address: loginEmail,
      password: loginPassword,
    };

    // クッキー文字列を構築
    var cookieString = Object.keys(cookieMap)
      .map(function (key) {
        return key + "=" + cookieMap[key];
      })
      .join("; ");

    var loginHeaders = {
      "User-Agent":
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
      Accept:
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
      "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
      "Content-Type": "application/x-www-form-urlencoded",
      Cookie: cookieString,
      Referer: loginUrl,
      Origin: "https://www.ecoauc.com",
    };

    var payloadString = Object.keys(loginPayload)
      .map(function (key) {
        return (
          encodeURIComponent(key) + "=" + encodeURIComponent(loginPayload[key])
        );
      })
      .join("&");

    console.log(
      "エコリング ログインペイロード:",
      payloadString.replace(loginPassword, "***")
    );

    var loginOptions = {
      method: "post",
      headers: loginHeaders,
      payload: payloadString,
      muteHttpExceptions: true,
      followRedirects: false,
    };

    // ログインURLを正しいエンドポイントに変更
    var loginPostUrl = "https://www.ecoauc.com/client/users/post-sign-in";
    var loginResponse = UrlFetchApp.fetch(loginPostUrl, loginOptions);
    console.log(
      "エコリング ログインレスポンスステータス:",
      loginResponse.getResponseCode()
    );

    // ログインエラーの場合、レスポンス内容を確認
    if (
      loginResponse.getResponseCode() !== 302 &&
      loginResponse.getResponseCode() !== 200
    ) {
      var loginErrorHtml = loginResponse.getContentText("UTF-8");
      console.log(
        "エコリング ログインエラーレスポンス（最初の1000文字）:",
        loginErrorHtml.substring(0, 1000)
      );
      logError_(
        "fetchEcoringHtml_",
        "ERROR",
        "ログインエラー",
        "ステータス: " +
          loginResponse.getResponseCode() +
          ", 内容: " +
          loginErrorHtml.substring(0, 500)
      );
    }

    // ログイン後のクッキーを更新
    var loginCookies = loginResponse.getAllHeaders()["Set-Cookie"];
    if (loginCookies) {
      if (Array.isArray(loginCookies)) {
        loginCookies.forEach(function (cookie) {
          var parts = cookie.split(";")[0].split("=");
          if (parts.length >= 2) {
            cookieMap[parts[0]] = parts.slice(1).join("=");
          }
        });
      } else {
        var parts = loginCookies.split(";")[0].split("=");
        if (parts.length >= 2) {
          cookieMap[parts[0]] = parts.slice(1).join("=");
        }
      }
    }

    // リダイレクト先を確認
    var locationHeader =
      loginResponse.getAllHeaders()["Location"] ||
      loginResponse.getAllHeaders()["location"];
    console.log("エコリング リダイレクト先:", locationHeader || "なし");

    // ステップ3: 目的のURLにアクセス（ログイン後のクッキーを使用）
    cookieString = Object.keys(cookieMap)
      .map(function (key) {
        return key + "=" + cookieMap[key];
      })
      .join("; ");

    var fetchOptions = {
      method: "get",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
        Accept:
          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        Cookie: cookieString,
        Referer: "https://www.ecoauc.com/client/market-prices",
      },
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(url, fetchOptions);
    var html = response.getContentText("UTF-8");

    if (response.getResponseCode() !== 200) {
      logError_(
        "fetchEcoringHtml_",
        "ERROR",
        "HTTPステータスエラー",
        "ステータスコード: " + response.getResponseCode()
      );
      throw new Error(
        "エコリングからのデータ取得に失敗しました。ステータスコード: " +
          response.getResponseCode()
      );
    }

    // ログインが必要なページにリダイレクトされていないかチェック
    if (
      html.indexOf("sign-in") > -1 &&
      html.indexOf("ログイン") > -1 &&
      html.indexOf("market-prices") === -1
    ) {
      console.log("エコリング HTMLレスポンスの一部:", html.substring(0, 500));
      logError_(
        "fetchEcoringHtml_",
        "ERROR",
        "ログイン失敗",
        "ログインページにリダイレクトされました。認証情報を確認してください。"
      );
      throw new Error(
        "エコリング ログインに失敗しました。認証情報を確認してください。"
      );
    }

    // 商品データが含まれているかチェック
    if (html.indexOf("show-case-title-block") === -1) {
      console.log(
        "エコリング 取得したHTMLに商品データが含まれていない可能性があります。"
      );
      console.log("エコリング HTMLの最初の1500文字:", html.substring(0, 1500));
      logWarning_(
        "fetchEcoringHtml_",
        "商品データなし",
        "HTMLに商品データが見つかりません。HTMLサイズ: " + html.length
      );
    }

    logInfo_("fetchEcoringHtml_", "HTML取得成功", "HTMLサイズ: " + html.length);
    return html;
  } catch (e) {
    logError_("fetchEcoringHtml_", "ERROR", "HTMLフェッチエラー", e.message);
    throw new Error("エコリング URLからのHTMLフェッチエラー: " + e.message);
  }
}

/**
 * エコリングのHTMLをパースして商品情報を抽出
 */
function parseEcoringFromHtml_(html) {
  var items = [];
  var H = String(html);
  var baseUrl = "https://www.ecoauc.com";

  console.log("エコリングパーサー - HTMLの長さ:", H.length);

  // 商品ブロックを抽出
  var blockRe =
    /<div\s+class="col-sm-6\s+col-md-4\s+col-lg-3"[^>]*>[\s\S]*?(?=<div\s+class="col-sm-6\s+col-md-4\s+col-lg-3"|<\/div>\s*<\/div>\s*<\/div>\s*<\/section>|$)/gi;
  var blocks = H.match(blockRe);

  if (!blocks || blocks.length === 0) {
    console.log("エコリング 商品ブロックが見つかりません");
    return items;
  }

  console.log("エコリング 商品ブロック数:", blocks.length);

  blocks.forEach(function (block, index) {
    // 詳細URL
    var detailUrl = "";
    var urlMatch = block.match(
      /<a[^>]+href="(\/client\/market-prices\/view\/[0-9]+)"[^>]*>/i
    );
    if (urlMatch) {
      detailUrl = baseUrl + htmlDecode_(urlMatch[1]);
    }

    // 画像URL
    var imageUrl = "";
    var imgMatch = block.match(
      /<img[^>]+class="[^"]*show-case-img-top[^"]*"[^>]+src="([^"]+)"/i
    );
    if (!imgMatch) {
      imgMatch = block.match(
        /<img[^>]+src="([^"]+)"[^>]+class="[^"]*show-case-img-top[^"]*"/i
      );
    }
    if (imgMatch) {
      imageUrl = htmlDecode_(imgMatch[1]);
      if (imageUrl && !imageUrl.startsWith("http")) {
        imageUrl = baseUrl + imageUrl;
      }
    }

    // 商品名
    var title = "";
    var titleMatch = block.match(
      /<div[^>]+class="[^"]*show-case-title-block[^"]*"[^>]*>[\s\S]*?<b>([^<]+)<\/b>/i
    );
    if (titleMatch) {
      title = htmlDecode_(titleMatch[1]).trim();
    }

    // 日付
    var date = "";
    var dateMatch = block.match(
      /<small[^>]+class="[^"]*show-case-daily[^"]*"[^>]*>([^<]+)<\/small>/i
    );
    if (dateMatch) {
      date = htmlDecode_(dateMatch[1]).trim();
    }

    // ランク
    var rank = "";
    var rankMatch = block.match(
      /<li[^>]+class="[^"]*canopy-2[^"]*"[^>]*>[\s\S]*?<span[^>]+class="[^"]*canopy-rank[^"]*"[^>]*>([^<]+)<\/span>/i
    );
    if (rankMatch) {
      var rawRank = htmlDecode_(rankMatch[1]).trim();
      rank = normalizeRank_(rawRank);
      // ランクNのデバッグ
      if (rawRank.toUpperCase() === "N" || rank === "N") {
        console.log(
          "エコリング - ランクN検出: raw='" +
            rawRank +
            "' → normalized='" +
            rank +
            "'"
        );
      }
    }

    // 価格
    var price = "";
    var priceAreaMatch = block.match(
      /<div[^>]+class="[^"]*item-text[^"]*"[^>]*>[\s\S]*?<span[^>]+class="[^"]*show-value[^"]*"[^>]*>([^<]+)<\/span>/i
    );
    if (priceAreaMatch) {
      var priceText = htmlDecode_(priceAreaMatch[1]).trim();
      price = normalizeNumber_(
        priceText.replace(/&yen;/g, "").replace(/[¥￥]/g, "")
      );
    }

    // ショップ
    var shop = "";
    var shopMatch = block.match(
      /<span[^>]+class="[^"]*market-title[^"]*"[^>]*>([\s\S]*?)<\/span>/i
    );
    if (shopMatch) {
      var shopHtml = shopMatch[1];
      shop = stripTags_(htmlDecode_(shopHtml)).trim();
      var shopNameMatch = shop.match(/【([^】]+)】/);
      if (shopNameMatch) {
        shop = shopNameMatch[1];
      }
    }

    if (detailUrl || imageUrl || title || date || rank || price) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: date,
          rank: rank,
          price: price,
          shop: shop,
          source: "エコリング",
          soldOut: false,
        })
      );
    }
  });

  console.log("エコリング パース完了:", items.length + "件");
  return items;
}
