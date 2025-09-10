/**
 * SBAのSAランク取得テスト
 */

// 実際のSBAスクレイパーでSAランクをテスト
function testSbaWithSARank() {
    console.log("=== SBA SAランク取得テスト開始 ===");
  
    // サンプルHTML（全角のSAランクを含む）
    var html = `
      <div class="p-auctions__list-item">
        <div class="p-product -auction">
          <a class="p-product__link" href="/auction/detail/12345">
            <div class="p-product__image" style="background-image: url('https://example.com/image1.jpg')"></div>
            <div class="p-product__detail">
              <div class="p-product__detail__title">エルメス ヴィドポッシュ</div>
              <ul class="p-product__detail__option">
                <li data-head="開催日"><strong>2025/01/10</strong></li>
                <li data-head="商品管理番号">12345</li>
                <li data-head="ランク" data-rank="ＳＡ"><strong>ＳＡ</strong></li>
                <li data-head="落札額"><strong>320,000yen</strong></li>
              </ul>
            </div>
          </a>
        </div>
      </div>
    `;
  
    // 実際のSBAスクレイパーを使用
    var items = parseSbaHtmlFile_(html);
  
    console.log("抽出されたアイテム数:", items.length);
  
    if (items.length > 0) {
      var item = items[0];
      console.log("\n=== アイテム詳細 ===");
      console.log("商品名:", item.title);
      console.log("ランク（生データ）:", item.rank);
      console.log("価格:", item.price);
  
      // normalizeRank_の動作確認
      console.log("\n=== normalizeRank_の処理 ===");
      var normalized = normalizeRank_(item.rank);
      console.log("正規化後のランク:", normalized);
      console.log("空文字列かどうか:", normalized === "");
  
      // normalizeRankForWrite_の動作確認
      console.log("\n=== normalizeRankForWrite_の処理 ===");
      var forWrite = normalizeRankForWrite_(item.rank);
      console.log("書き込み用に正規化後:", forWrite);
      console.log("空文字列かどうか:", forWrite === "");
    }
  
    console.log("\n=== 有効なランクリスト ===");
    var validRanks = ["N", "S", "SA", "A", "AB", "B", "BC", "C"];
    console.log("有効なランク:", validRanks.join(", "));
    console.log("'SA'が含まれているか:", validRanks.includes("SA"));
  
    console.log("\n=== SBA SAランク取得テスト終了 ===");
  }
  
  // 全てのランクタイプをテスト
  function testAllRankTypes() {
    console.log("=== 全ランクタイプテスト開始 ===");
  
    var testRanks = ["Ｎ", "Ｓ", "ＳＡ", "Ａ", "ＡＢ", "Ｂ", "ＢＣ", "Ｃ"];
  
    testRanks.forEach(function (rank) {
      console.log("\n全角ランク '" + rank + "' のテスト:");
  
      // 全角を半角に変換
      var halfWidth = rank.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
      });
      console.log("  半角変換後:", halfWidth);
  
      // normalizeRank_でテスト
      var normalized = normalizeRank_(rank);
      console.log("  normalizeRank_結果:", normalized || "(空文字列)");
  
      // normalizeRankForWrite_でテスト
      var forWrite = normalizeRankForWrite_(rank);
      console.log("  normalizeRankForWrite_結果:", forWrite || "(空文字列)");
    });
  
    console.log("\n=== 全ランクタイプテスト終了 ===");
  }
  
  // 実際のHTMLファイルからランクを確認
  function checkRanksInActualSbaHtml() {
    console.log("=== 実際のSBA HTMLファイルのランク確認 ===");
  
    // 実際のファイルを読み込む場合は、別途ファイル読み込み処理が必要
    // ここではテスト用に直接HTMLを指定
    var sampleHtml = `
      <span class="icon" data-rank="ＳＡ"> ＳＡ </span>
      <span class="icon" data-rank="Ａ"> Ａ </span>
      <span class="icon" data-rank="ＡＢ"> ＡＢ </span>
    `;
  
    // data-rank属性の値を抽出
    var matches = sampleHtml.match(/data-rank="([^"]+)"/g);
    if (matches) {
      console.log("\n見つかったdata-rank属性:");
      matches.forEach(function (match) {
        var rank = match.match(/data-rank="([^"]+)"/)[1];
        console.log("  元の値: '" + rank + "'");
        console.log("  正規化後: '" + normalizeRank_(rank) + "'");
      });
    }
  
    console.log("\n=== 実際のSBA HTMLファイルのランク確認終了 ===");
  }
  