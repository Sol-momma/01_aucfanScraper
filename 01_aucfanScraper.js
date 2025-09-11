/**
 * オークファン専用スクレイパー
 * オークファンからの商品情報取得処理
 */

// ===== オークファン スクレイピング =====

// パフォーマンス最適化：事前コンパイル済み正規表現
const AUC_PRICE_REGEX = {
  RAKUSATSU_SPAN: /<li class="price"[^>]*>\s*<span[^>]*>落札<\/span>\s*([^<]+)/i,
  RAKUSATSU_TEXT: /落札価格[:：]?\s*([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i,
  END_PRICE: /終了価格[:：]?\s*([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i,
  PRICE_LI: /<li class="price"[^>]*>([\s\S]*?)<\/li>/i,
  DATA_PRICE: /data-price="([0-9,]+)"/i,
  PRICE_VALUE: /class="[^"]*price__value[^"]*"[^>]*>\s*([^<]+)/i,
  YEN_NUM: /([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i,
  YEN_MARK: /¥\s*([0-9]{1,3}(?:,[0-9]{3})*)/i,
  LEADING_TEXT: /^\s*([^<]+)/i
};

const AUC_EXCLUDED_SET = new Set([3980, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000]);

const AUC_RANK_REGEX = {
  YAHOO_STATE: /商品状態[：:\s]*([^<\n]+)/i,
  MERC_PATTERNS: [
    /新品、未使用/i, /未使用に近い/i, /目立った傷や汚れなし/i,
    /やや傷や汚れあり/i, /傷や汚れあり/i, /全体的に状態が悪い/i
  ]
};

const AUC_COMMON_REGEX = {
  ITEM: /<li class="item">\s*<ul>([\s\S]*?)<\/ul>\s*<\/li>/gi,
  DETAIL_PATH: /<li class="title">[\s\S]*?<a\s+href="([^"]+)"/i,
  IMAGE: /<img[^>]+class="thumbnail"[^>]+src="([^"]+)"/i,
  TITLE_ANCHOR: /<li class="itemName">[\s\S]*?<a[^>]*>([\s\S]*?)<\/a>/i,
  TITLE_FALLBACK: /<li class="itemName">([\s\S]*?)<\/li>/i,
  END: /<li class="end">\s*([^<]+)/i
};

/**
 * オークファンのHTMLを取得
 */
function fetchAucfanHtml_(url) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'auc_html_' + Utilities.base64EncodeWebSafe(url).slice(0, 64);
    var cached = cache.get(cacheKey);
    if (cached) {
      return cached;
    }
    var options = getCommonHttpOptions_();
    var response = UrlFetchApp.fetch(url, options);
    validateHttpResponse_(response, "オークファン");
    var text = getResponseTextWithBestCharset_(response);
    if (text && text.length <= 90000) {
      cache.put(cacheKey, text, 300); // 5分キャッシュ
    } else {
      console.warn("HTMLが大きいためキャッシュをスキップ:", text ? text.length + "文字" : 0);
    }
    return text;
  } catch (e) {
    throw new Error("オークファン URLからのHTMLフェッチエラー: " + e.message);
  }
}

/**
 * オークファンデータを取得（HTMLまたはURLから）
 */
function getAucfanDataFromSheetOrUrl_(sheet, urlRow, urlCol) {
  var t0_total = Date.now(); // 関数全体の開始時間
  console.log("--- getAucfanDataFromSheetOrUrl_ 全体計測開始 ---");

  try {
    // まず15行目以降にHTMLがあるかチェック
    var t0_readSheet = Date.now();
    var htmlFromSheet = readHtmlFromRow_(sheet, 15);
    var t1_readSheet = Date.now();
    console.log("  HTMLをシートから読み込み (readHtmlFromRow_):", (t1_readSheet - t0_readSheet) + "ms");

    if (htmlFromSheet) {
      // HTMLソースがオークファンかどうかを判定
      var source = detectSource_(htmlFromSheet);
      if (source === "オークファン" || source === "aucfan") {
        console.log("  67行目以降のHTMLからオークファンデータを取得します");
        var t0_parse = Date.now();
        var items = parseAucfanFromHtml_(htmlFromSheet);
        var t1_parse = Date.now();
        console.log("  HTMLパース (parseAucfanFromHtml_):", (t1_parse - t0_parse) + "ms");
        console.log("  オークファンデータを取得しました:", items.length + "件");
        
        var t1_total = Date.now();
        console.log("--- getAucfanDataFromSheetOrUrl_ 全体計測終了 ---");
        console.log("  関数全体の処理時間:", (t1_total - t0_total) + "ms");
        return items;
      } else {
        console.log("  67行目以降のHTMLはオークファンではありません:", source);
      }
    }

    // HTMLがない場合はURLから取得を試行
    var url = sheet.getRange(urlRow || 110, urlCol || 2).getValue();
    if (url && url.toString().trim()) {
      var urlStr = url.toString().trim();
      if (urlStr.startsWith("http")) {
        console.log("  URLからオークファンデータを取得します:", urlStr);
        var t0_fetchHtml = Date.now();
        var html = fetchAucfanHtml_(urlStr);
        var t1_fetchHtml = Date.now();
        console.log("  URLからHTMLをフェッチ (fetchAucfanHtml_):", (t1_fetchHtml - t0_fetchHtml) + "ms");

        var t0_parse = Date.now();
        var items = parseAucfanFromHtml_(html);
        var t1_parse = Date.now();
        console.log("  HTMLパース (parseAucfanFromHtml_):", (t1_parse - t0_parse) + "ms");
        console.log("  オークファンデータを取得しました:", items.length + "件");

        var t1_total = Date.now();
        console.log("--- getAucfanDataFromSheetOrUrl_ 全体計測終了 ---");
        console.log("  関数全体の処理時間:", (t1_total - t0_total) + "ms");
        return items;
      }
    }

    console.log("  オークファンのHTMLもURLも見つかりません");
    var t1_total = Date.now();
    console.log("--- getAucfanDataFromSheetOrUrl_ 全体計測終了 ---");
    console.log("  関数全体の処理時間:", (t1_total - t0_total) + "ms");
    return [];
  } catch (e) {
    console.warn("オークファンデータの取得に失敗:", e.message);
    var t1_total = Date.now();
    console.log("--- getAucfanDataFromSheetOrUrl_ 全体計測終了 (エラー発生) ---");
    console.log("  関数全体の処理時間:", (t1_total - t0_total) + "ms");
    return [];
  }
}


// ===== 商品ランク機能（Aucfan専用・プレフィックス付きで衝突回避） =====

/**
 * Phase 1: 商品状態のテキストを抽出（最適化版）
 * @param {string} htmlBlock 商品ブロックのHTML
 * @returns {string} 商品状態テキスト（例: "中古", "目立った傷や汚れなし"）
 */
function auc_extractConditionText_(htmlBlock) {
  const m = htmlBlock.match(AUC_RANK_REGEX.YAHOO_STATE);
  if (m) return m[1].trim();
  
  for (let i = 0; i < AUC_RANK_REGEX.MERC_PATTERNS.length; i++) {
    const match = htmlBlock.match(AUC_RANK_REGEX.MERC_PATTERNS[i]);
    if (match) return match[0];
  }
  return "";
}

/**
 * Phase 2: サイト種別を判定
 * @param {string} htmlBlock 商品ブロックのHTML
 * @returns {string} サイト（"yahoo" | "mercari"）
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

/** Phase 3: ヤフオク商品状態→ランク マッピング */
const AUC_YAHOO_CONDITION_RANK_MAP = {
  '新品': 'S',
  '未使用': 'S', 
  '新品未使用': 'S',
  '未使用未開封': 'S',
  '新同': 'SA',
  '未使用に近い': 'SA',
  '極美品': 'SA',
  '美品': 'A',
  '目立った傷や汚れなし': 'A',
  'やや傷や汚れあり': 'B',
  '傷や汚れあり': 'C',
  '中古': 'B',  // デフォルト
  '全体的に状態が悪い': 'D',
  '訳あり': 'C',
  'ジャンク': 'D'
};

/** Phase 3: メルカリ商品状態→ランク マッピング */
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
 * @param {string} conditionText 商品状態テキスト
 * @returns {string} ランク（S|SA|A|B|C|D）
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
 * @param {string} conditionText 商品状態テキスト
 * @returns {string} ランク（S|SA|A|B|C|D）
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
 * @param {string} conditionText 商品状態テキスト
 * @param {string} siteType サイト種別
 * @returns {string} ランク
 */
function auc_convertConditionToRank_(conditionText, siteType) {
  if (!conditionText) return "B"; // デフォルトはB
  
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
  var t0 = Date.now();
  const items = [];
  const base = "https://pro.aucfan.com";
  
  AUC_COMMON_REGEX.ITEM.lastIndex = 0;
  let m;
  var itemCount = 0;
  var timeBreakdown = { detail: 0, image: 0, title: 0, price: 0, end: 0, rank: 0, create: 0 };

  while ((m = AUC_COMMON_REGEX.ITEM.exec(html)) !== null) {
    itemCount++;
    const block = m[1];

    // 詳細URL
    var t1 = Date.now();
    const detailPath = firstMatch_(block, AUC_COMMON_REGEX.DETAIL_PATH);
    const detailUrl = detailPath
      ? (detailPath.startsWith("http") ? detailPath : base + detailPath)
      : "";
    timeBreakdown.detail += Date.now() - t1;

    // 画像URL
    t1 = Date.now();
    const imageUrl = firstMatch_(block, AUC_COMMON_REGEX.IMAGE);
    timeBreakdown.image += Date.now() - t1;

    // タイトル抽出（軽量化版）
    t1 = Date.now();
    let title = "";
    const titleRaw = firstMatch_(block, AUC_COMMON_REGEX.TITLE_ANCHOR) ||
                     firstMatch_(block, AUC_COMMON_REGEX.TITLE_FALLBACK);

    if (titleRaw) {
      // htmlDecode → stripTags → 正規化を1回で
      title = titleRaw
        .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, " ")
        .replace(/<[^>]*>/g, "")
        .replace(/\s+/g, " ")
        .trim();
    }
    timeBreakdown.title += Date.now() - t1;

    // 価格抽出（最優先＋軽量化版）
    t1 = Date.now();
    let price = "";

    // 落札価格を最優先（見つかったら即終了）
    let endPrice = firstMatch_(block, AUC_PRICE_REGEX.RAKUSATSU_SPAN) ||
                   firstMatch_(block, AUC_PRICE_REGEX.RAKUSATSU_TEXT) ||
                   firstMatch_(block, AUC_PRICE_REGEX.END_PRICE);

    if (endPrice) {
      const normalized = normalizeNumber_(endPrice);
      if (normalized && normalized !== "0") {
        price = normalized;
      }
    }

    // 落札価格が見つからない場合のみ、軽量候補探索
    if (!price) {
      const priceLi = firstMatch_(block, AUC_PRICE_REGEX.PRICE_LI);
      const area = priceLi || block;
      let min = Number.POSITIVE_INFINITY;

      // 軽い順に候補をチェック（見つかり次第最小更新）
      [
        firstMatch_(area, AUC_PRICE_REGEX.DATA_PRICE),
        firstMatch_(area, AUC_PRICE_REGEX.PRICE_VALUE),
        firstMatch_(area, AUC_PRICE_REGEX.YEN_NUM),
        firstMatch_(area, AUC_PRICE_REGEX.YEN_MARK),
        priceLi ? firstMatch_(area, AUC_PRICE_REGEX.LEADING_TEXT) : null
      ].forEach(function(candidate) {
        if (!candidate) return;
        const nStr = normalizeNumber_(candidate);
        if (!nStr || nStr === "0") return;
        const n = parseInt(nStr, 10);
        if (isNaN(n) || n < 100 || AUC_EXCLUDED_SET.has(n)) return;
        if (n < min) min = n;
      });

      if (isFinite(min)) price = String(min);
    }
    timeBreakdown.price += Date.now() - t1;

    // 終了日
    t1 = Date.now();
    const endTxt = firstMatch_(block, AUC_COMMON_REGEX.END);
    timeBreakdown.end += Date.now() - t1;

    // 商品状態抽出とランク変換
    t1 = Date.now();
    const conditionText = auc_extractConditionText_(block);
    const siteType = auc_detectSiteType_(block);
    const rank = auc_convertConditionToRank_(conditionText, siteType) || "B";
    timeBreakdown.rank += Date.now() - t1;

    t1 = Date.now();
    if (detailUrl || imageUrl || price || endTxt || title) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: endTxt || "",
          rank: rank,
          price: price || "",
          shop: "",
          source: "オークファン",
          soldOut: false,
        })
      );
    }
    timeBreakdown.create += Date.now() - t1;
  }

  var totalTime = Date.now() - t0;
  console.log("=== parseAucfanFromHtml_ パフォーマンス分析（最適化後） ===");
  console.log("総処理時間:", totalTime + "ms");
  console.log("処理アイテム数:", itemCount + "件");
  console.log("1件あたり平均:", Math.round(totalTime / Math.max(itemCount, 1)) + "ms");
  console.log("時間内訳:");
  console.log("  詳細URL:", timeBreakdown.detail + "ms (" + Math.round(timeBreakdown.detail / totalTime * 100) + "%)");
  console.log("  画像URL:", timeBreakdown.image + "ms (" + Math.round(timeBreakdown.image / totalTime * 100) + "%)");
  console.log("  タイトル:", timeBreakdown.title + "ms (" + Math.round(timeBreakdown.title / totalTime * 100) + "%)");
  console.log("  価格抽出:", timeBreakdown.price + "ms (" + Math.round(timeBreakdown.price / totalTime * 100) + "%)");
  console.log("  終了日:", timeBreakdown.end + "ms (" + Math.round(timeBreakdown.end / totalTime * 100) + "%)");
  console.log("  ランク:", timeBreakdown.rank + "ms (" + Math.round(timeBreakdown.rank / totalTime * 100) + "%)");
  console.log("  オブジェクト作成:", timeBreakdown.create + "ms (" + Math.round(timeBreakdown.create / totalTime * 100) + "%)");

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
    
    { conditionText: '', siteType: 'yahoo', expected: 'B' },
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
 * オークファンHTML読み取りテスト関数
 */
function testAucfanHtmlFromSheet() {
  console.log("=== オークファンHTML読み取りテスト (getAucfanDataFromSheetOrUrl_経由) ===");

  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    
    // getAucfanDataFromSheetOrUrl_ を呼び出して全体のフローをテスト
    var items = getAucfanDataFromSheetOrUrl_(sheet); 

    if (items && items.length > 0) {
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
        console.log("  ランク:", items[i].rank || "なし");
      }
      console.log("\n🎉 テスト成功！");
    } else {
      console.log("❌ 商品データが取得できませんでした。ログを確認してください。");
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
  console.log("===== ランク機能 統合テスト =====");
  
  console.log("\nPhase 1: 商品状態抽出テスト");
  testAucfanExtractConditionText();
  
  console.log("\nPhase 2: サイト判定テスト");
  testAucfanDetectSiteType();
  
  console.log("\nPhase 3: ランク変換テスト");
  testAucfanConvertConditionToRank();
  
  console.log("\n全てのPhaseのテストが完了しました。");
  console.log("実データでの確認: testAucfanHtmlFromSheet() を実行。");
}