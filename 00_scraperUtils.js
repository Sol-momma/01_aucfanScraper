/**
 * スクレイピング共通ユーティリティ
 * 各サイトのスクレイパーで共通して使用される関数群
 */

// ===== 共通ユーティリティ関数 =====

/**
 * HTMLエンティティをデコード
 */
function htmlDecode_(s) {
    if (!s) return "";
    return s
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&nbsp;/g, " ");
  }
  
  /**
   * 文字エンコーディングを検出してレスポンステキストを取得
   */
  function detectCharsetFromResponse_(response) {
    var contentType = response.getHeaders()["Content-Type"] || "";
    var charsetMatch = contentType.match(/charset=([^;]+)/i);
    if (charsetMatch) {
      return charsetMatch[1].trim();
    }
  
    var html = response.getContentText();
    var metaCharsetMatch = html.match(/<meta[^>]+charset=["']?([^"'>]+)/i);
    if (metaCharsetMatch) {
      return metaCharsetMatch[1].trim();
    }
  
    return "UTF-8";
  }
  
  /**
   * 最適な文字エンコーディングでレスポンステキストを取得
   */
  function getResponseTextWithBestCharset_(response) {
    var charset = detectCharsetFromResponse_(response);
    try {
      return response.getContentText(charset);
    } catch (e) {
      return response.getContentText("UTF-8");
    }
  }
  
  /**
   * HTMLタグを除去
   */
  function stripTags_(s) {
    if (!s) return "";
    return s.replace(/<[^>]*>/g, "");
  }
  
  /**
   * 数値文字列を正規化
   */
  function normalizeNumber_(s) {
    if (!s) return "";
    var cleaned = String(s).replace(/[^\d]/g, "");
    return cleaned || "";
  }
  
  /**
   * 正規表現の最初のマッチを取得
   */
  function firstMatch_(text, re) {
    var match = text.match(re);
    return match ? match[1] : "";
  }
  
  /**
   * ランクを正規化
   */
  function normalizeRank_(rank) {
    if (!rank) return "";
  
    // 全角を半角に変換
    var normalized = rank.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
    });
  
    // 大文字に変換
    var upperCase = normalized.toUpperCase();
  
    // デバッグログ（SAの場合のみ）
    if (rank === "ＳＡ" || upperCase === "SA") {
      console.log("normalizeRank_デバッグ:");
      console.log("  入力:", rank);
      console.log("  半角変換後:", normalized);
      console.log("  大文字変換後:", upperCase);
    }
  
    var validRanks = ["N", "S", "SA", "A", "AB", "B", "BC", "C"];
    var isValid = validRanks.includes(upperCase);
  
    if (rank === "ＳＡ" || upperCase === "SA") {
      console.log("  有効チェック:", isValid);
    }
  
    return isValid ? upperCase : "";
  }
  
  /**
   * 共通のHTTPリクエストオプションを生成
   */
  function getCommonHttpOptions_() {
    return {
      method: "get",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
      },
      muteHttpExceptions: true,
    };
  }
  
  /**
   * HTTPレスポンスの共通検証
   */
  function validateHttpResponse_(response, siteName) {
    if (response.getResponseCode() !== 200) {
      throw new Error(
        siteName +
          "からのHTMLの取得に失敗しました。ステータスコード: " +
          response.getResponseCode()
      );
    }
  }
  
  /**
   * 商品データの標準フォーマットを作成
   */
  function createItemData_(data) {
    return {
      title: data.title || "",
      detailUrl: data.detailUrl || "",
      imageUrl: data.imageUrl || "",
      date: data.date || "",
      rank: data.rank || "",
      price: data.price || "",
      shop: data.shop || "",
      source: data.source || "",
      soldOut: data.soldOut === true ? true : false, // 明示的にbooleanに変換
    };
  }
  
  /**
   * HTMLソースを自動判定
   */
  function detectSource_(html) {
    if (html.indexOf("rakuten.co.jp") > -1 || html.indexOf("楽天市場") > -1) {
      return "楽天";
    } else if (html.indexOf("starbuyers-global-auction.com") > -1) {
      return "SBA";
    } else if (html.indexOf("aucfan.com") > -1) {
      return "オークファン";
    } else if (html.indexOf("ecoring.com") > -1) {
      return "エコリング";
    } else if (html.indexOf("yahoo.co.jp") > -1) {
      return "ヤフオク";
    }
    return "不明";
  }
  
  /**
   * 指定行以降からHTMLを読み取り
   */
  function readHtmlFromRow_(sheet, startRow) {
    if (!startRow) startRow = 67; // デフォルトは67行目から
  
    var lastRow = sheet.getLastRow();
    if (lastRow < startRow) {
      console.log("B" + startRow + "行目以降にHTMLがありません");
      return "";
    }
  
    var vals = sheet
      .getRange(startRow, 2, lastRow - startRow + 1, 1) // B列から取得
      .getValues()
      .map(function (r) {
        return r[0];
      });
  
    var joined = vals
      .filter(function (v) {
        return v !== "" && v !== null && v !== undefined;
      })
      .join("\n");
  
    if (!joined.trim()) {
      console.log("B" + startRow + "行目以降にHTMLデータが見つかりません");
      return "";
    }
  
    console.log(
      "HTML読み取り完了: " + joined.length + "文字 (B" + startRow + "行目以降)"
    );
    return joined.trim();
  }
  