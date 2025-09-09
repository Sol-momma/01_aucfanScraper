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


// ===== 商品ランク機能 =====

/**
 * Phase 1: 商品状態のテキストを抽出する関数
 * @param {string} htmlBlock - 商品ブロックのHTML
 * @return {string} 商品状態のテキスト（例: "中古", "目立った傷や汚れなし"）
 */
function auc_extractConditionText_(htmlBlock) {
  // ヤフオクパターン: "商品状態 中古" のような形式
  const yahooPattern = /商品状態[：:\s]*([^<\n]+)/i;
  const yahooMatch = htmlBlock.match(yahooPattern);
  if (yahooMatch) {
    return yahooMatch[1].trim();
  }
  
  // メルカリパターン: 状態を直接記載
  const mercariPatterns = [
    /新品、未使用/i,
    /未使用に近い/i,
    /目立った傷や汚れなし/i,
    /やや傷や汚れあり/i,
    /傷や汚れあり/i,
    /全体的に状態が悪い/i
  ];
  
  for (let pattern of mercariPatterns) {
    const match = htmlBlock.match(pattern);
    if (match) {
      return match[0];
    }
  }
  
  // 何も見つからない場合
  return "";
}

/**
 * Phase 2: サイト種別を判定する関数  
 * @param {string} htmlBlock - 商品ブロックのHTML
 * @return {string} サイト名（"yahoo" または "mercari"）
 */
function auc_detectSiteType_(htmlBlock) {
  // ヤフオクの特徴を探す
  if (htmlBlock.includes('yahoo') || 
      htmlBlock.includes('aucfan') || 
      htmlBlock.includes('商品状態') ||
      htmlBlock.includes('ヤフオク')) {
    return "yahoo";
  }
  
  // メルカリの特徴を探す
  if (htmlBlock.includes('mercari') || 
      htmlBlock.includes('メルカリ') ||
      htmlBlock.includes('フリマ')) {
    return "mercari";
  }
  
  // デフォルト（分からない場合はヤフオクとして扱う）
  return "yahoo";
}

// ===== Phase 3: ランク変換機能 =====

/**
 * Phase 3: ヤフオク商品状態とランクのマッピング定義
 */
const AUC_YAHOO_CONDITION_RANK_MAP = {
  '新品': 'S',
  '未使用': 'S', 
  '未使用に近い': 'SA',
  '目立った傷や汚れなし': 'A',
  'やや傷や汚れあり': 'B',
  '傷や汚れあり': 'C',
  '中古': 'B',  // デフォルト
  '全体的に状態が悪い': 'D'
};

/**
 * Phase 3: メルカリ商品状態とランクのマッピング定義
 */
const AUC_MERCARI_CONDITION_RANK_MAP = {
  '新品、未使用': 'S',
  '未使用に近い': 'SA',
  '目立った傷や汚れなし': 'A', 
  'やや傷や汚れあり': 'B',
  '傷や汚れあり': 'C',
  '全体的に状態が悪い': 'D'
};

/**
 * Phase 3: ヤフオクの商品状態をランクに変換
 * @param {string} conditionText - 商品状態テキスト
 * @return {string} ランク（S, SA, A, B, C, D）
 */
function auc_convertYahooConditionToRank_(conditionText) {
  if (!conditionText) return "";
  
  // 完全一致を試す
  if (AUC_YAHOO_CONDITION_RANK_MAP[conditionText]) {
    return AUC_YAHOO_CONDITION_RANK_MAP[conditionText];
  }
  
  // 部分一致を試す
  for (let condition in AUC_YAHOO_CONDITION_RANK_MAP) {
    if (conditionText.includes(condition)) {
      return AUC_YAHOO_CONDITION_RANK_MAP[condition];
    }
  }
  
  // 何も一致しない場合はデフォルト
  return "B";
}

/**
 * Phase 3: メルカリの商品状態をランクに変換
 * @param {string} conditionText - 商品状態テキスト  
 * @return {string} ランク（S, SA, A, B, C, D）
 */
function auc_convertMercariConditionToRank_(conditionText) {
  if (!conditionText) return "";
  
  // 完全一致を試す
  if (AUC_MERCARI_CONDITION_RANK_MAP[conditionText]) {
    return AUC_MERCARI_CONDITION_RANK_MAP[conditionText];
  }
  
  // 部分一致を試す
  for (let condition in AUC_MERCARI_CONDITION_RANK_MAP) {
    if (conditionText.includes(condition)) {
      return AUC_MERCARI_CONDITION_RANK_MAP[condition];
    }
  }
  
  // 何も一致しない場合はデフォルト
  return "B";
}

/**
 * Phase 3: 統一ランク変換関数（メイン）
 * @param {string} conditionText - 商品状態テキスト
 * @param {string} siteType - サイト種別
 * @return {string} ランク
 */
function auc_convertConditionToRank_(conditionText, siteType) {
  if (!conditionText) return "";
  
  let map;
  if (siteType === "mercari") {
    map = AUC_MERCARI_CONDITION_RANK_MAP;
  } else {
    map = AUC_YAHOO_CONDITION_RANK_MAP;
  }
  
  // 完全一致を試す
  if (map[conditionText]) {
    return map[conditionText];
  }
  
  // 部分一致を試す
  for (let condition in map) {
    if (conditionText.includes(condition)) {
      return map[condition];
    }
  }
  
  // 何も一致しない場合はデフォルト
  return "B";
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

    // 🆕 商品状態抽出とランク変換
    const conditionText = auc_extractConditionText_(block);
    const siteType = auc_detectSiteType_(block);  
    const rank = auc_convertConditionToRank_(conditionText, siteType);

    if (detailUrl || imageUrl || price || endTxt || title) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: endTxt || "",
          rank: rank, // 🆕 ここが変わる！
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

// ===== テスト関数 =====

/**
 * Phase 1のテスト関数
 */
function testAucfanExtractConditionText() {
  console.log("=== Phase 1: 商品状態抽出テスト ===");
  
  // テストケース
  const testCases = [
    { html: '商品状態 新品', expected: '新品' },
    { html: '商品状態: 中古', expected: '中古' },
    { html: '商品状態　未使用に近い', expected: '未使用に近い' },
    { html: '目立った傷や汚れなし', expected: '目立った傷や汚れなし' },
    { html: 'やや傷や汚れあり', expected: 'やや傷や汚れあり' },
    { html: '何も状態情報がないHTML', expected: '' }
  ];
  
  for (let i = 0; i < testCases.length; i++) {
    const testCase = testCases[i];
    const result = auc_extractConditionText_(testCase.html);
    
    console.log(`テスト${i + 1}:`);
    console.log(`  入力: "${testCase.html}"`);
    console.log(`  結果: "${result}"`);
    console.log(`  期待値: "${testCase.expected}"`);
    console.log(`  判定: ${result === testCase.expected ? '✅ 成功' : '❌ 失敗'}`);
    console.log('---');
  }
}

/**
 * Phase 2のテスト関数
 */
function testAucfanDetectSiteType() {
  console.log("=== Phase 2: サイト判定テスト ===");
  
  // テストケース
  const testCases = [
    { html: '商品状態 新品 yahoo', expected: 'yahoo' },
    { html: 'aucfan.com の商品', expected: 'yahoo' },
    { html: 'メルカリで販売', expected: 'mercari' },
    { html: 'mercari フリマアプリ', expected: 'mercari' },
    { html: '普通のHTMLテキスト', expected: 'yahoo' } // デフォルト
  ];
  
  for (let i = 0; i < testCases.length; i++) {
    const testCase = testCases[i];
    const result = auc_detectSiteType_(testCase.html);
    
    console.log(`テスト${i + 1}:`);
    console.log(`  入力: "${testCase.html}"`);
    console.log(`  結果: "${result}"`);
    console.log(`  期待値: "${testCase.expected}"`);
    console.log(`  判定: ${result === testCase.expected ? '✅ 成功' : '❌ 失敗'}`);
    console.log('---');
  }
}

/**
 * Phase 3のテスト関数
 */
function testAucfanConvertConditionToRank() {
  console.log("=== Phase 3: ランク変換テスト ===");
  
  // テストケース
  const testCases = [
    { conditionText: '新品', siteType: 'yahoo', expected: 'S' },
    { conditionText: '中古', siteType: 'yahoo', expected: 'B' },
    { conditionText: '未使用に近い', siteType: 'yahoo', expected: 'SA' },
    { conditionText: '目立った傷や汚れなし', siteType: 'yahoo', expected: 'A' },
    
    { conditionText: '新品、未使用', siteType: 'mercari', expected: 'S' },
    { conditionText: '未使用に近い', siteType: 'mercari', expected: 'SA' },
    { conditionText: '目立った傷や汚れなし', siteType: 'mercari', expected: 'A' },
    { conditionText: 'やや傷や汚れあり', siteType: 'mercari', expected: 'B' },
    
    { conditionText: '', siteType: 'yahoo', expected: '' },
    { conditionText: '不明な状態', siteType: 'yahoo', expected: 'B' }
  ];
  
  for (let i = 0; i < testCases.length; i++) {
    const testCase = testCases[i];
    const result = auc_convertConditionToRank_(testCase.conditionText, testCase.siteType);
    
    console.log(`テスト${i + 1}:`);
    console.log(`  商品状態: "${testCase.conditionText}"`);
    console.log(`  サイト: "${testCase.siteType}"`);
    console.log(`  結果: "${result}"`);
    console.log(`  期待値: "${testCase.expected}"`);
    console.log(`  判定: ${result === testCase.expected ? '✅ 成功' : '❌ 失敗'}`);
    console.log('---');
  }
}

/**
 * デバッグ用テスト関数
 */
function debugRankConversion() {
  console.log("=== デバッグ：ランク変換テスト ===");
  
  // 直接マッピングをテスト
  console.log("1. マッピング定義の確認:");
  console.log("AUC_YAHOO_CONDITION_RANK_MAP['新品']:", AUC_YAHOO_CONDITION_RANK_MAP['新品']);
  console.log("AUC_YAHOO_CONDITION_RANK_MAP['中古']:", AUC_YAHOO_CONDITION_RANK_MAP['中古']);
  
  // 関数を直接テスト
  console.log("\n2. Yahoo変換関数のテスト:");
  const yahooResult1 = auc_convertYahooConditionToRank_('新品');
  console.log("convertYahooConditionToRank_('新品'):", yahooResult1);
  
  const yahooResult2 = auc_convertYahooConditionToRank_('中古');
  console.log("convertYahooConditionToRank_('中古'):", yahooResult2);
  
  // 統一関数をテスト
  console.log("\n3. 統一変換関数のテスト:");
  const unifiedResult1 = auc_convertConditionToRank_('新品', 'yahoo');
  console.log("convertConditionToRank_('新品', 'yahoo'):", unifiedResult1);
  
  const unifiedResult2 = auc_convertConditionToRank_('中古', 'yahoo');
  console.log("convertConditionToRank_('中古', 'yahoo'):", unifiedResult2);
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
      console.log("  ランク:", items[i].rank || "なし"); // 🆕 ランク表示を追加
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

/**
 * 🎯 統合テスト：全Phase機能の動作確認
 */
function testAucfanAllRankFeatures() {
  console.log("🎯 ===== ランク機能 統合テスト ===== 🎯");
  
  console.log("\n📋 Phase 1: 商品状態抽出テスト");
  testAucfanExtractConditionText();
  
  console.log("\n📋 Phase 2: サイト判定テスト");
  testAucfanDetectSiteType();
  
  console.log("\n📋 Phase 3: ランク変換テスト");
  testAucfanConvertConditionToRank();
  
  console.log("\n🎉 全てのPhaseのテストが完了しました！");
  console.log("実際のデータでテストするには testAucfanHtmlFromSheet() を実行してください。");
}