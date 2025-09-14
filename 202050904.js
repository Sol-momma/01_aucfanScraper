// /** ================== 定数 ================== **/
// const COL_B = 2; // B列
// const COL_C = 3; // C列（出力起点）
// const ROW_HTML_INPUT = 174; // HTML貼付開始行（ヤフオクセクション後）
// const ROW_HTML_START = ROW_HTML_INPUT; // 互換性のため残す
// const ROW_LABELS_START = 10; // ラベル群の開始行（目安）
// const ROW_CLEAR_10 = 10; // 10行目をクリア対象
// const ROW_CLEAR_11 = 11; // 11行目をクリア対象
// const ROW_CLEAR_13 = 13; // 13行目をクリア対象（チェックボックス）
// const ROW_CLEAR_BLOCK_FROM = 14; // 14行目から
// const ROW_CLEAR_BLOCK_TO = 21; // 21行目までをクリア対象（ショップを含む）
// const ROW_CLEAR_22 = 22; // 22行目をクリア対象
// const ROW_CLEAR_23 = 23; // 23行目をクリア対象
// const MAX_OUTPUT_ITEMS = 20; // C〜V（20列）に合わせる場合
// const LABELS = [
//   "詳細URL",
//   "画像URL",
//   "日付",
//   "ランク",
//   "価格",
//   "ショップ",
//   "引用サイト",
// ]; // B10〜B17で使う想定のラベル
// const DO_AUTO_RESIZE = false; // ← ここを false にしておく
// const APPLY_CLIP_WRAP = true; // ← 文字折返しで高さ変化を防ぐなら true
// const MAX_OUTPUT_COLS = 20; // C～V

// /** ============== エラーログ機能 ============== **/

// // エラーログシートを取得または作成
// function getOrCreateErrorLogSheet_() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var errorSheet = ss.getSheetByName("エラーログ");
  
//   if (!errorSheet) {
//     // 現在のアクティブシートを記憶
//     var currentSheet = ss.getActiveSheet();
    
//     // エラーログシートが存在しない場合は作成
//     errorSheet = ss.insertSheet("エラーログ");
    
//     // ヘッダーを設定
//     var headers = ["タイムスタンプ", "関数名", "エラータイプ", "エラーメッセージ", "詳細情報"];
//     errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
//     // ヘッダーのスタイルを設定
//     errorSheet.getRange(1, 1, 1, headers.length)
//       .setBackground("#4285F4")
//       .setFontColor("#FFFFFF")
//       .setFontWeight("bold");
    
//     // 列幅を調整
//     errorSheet.setColumnWidth(1, 150); // タイムスタンプ
//     errorSheet.setColumnWidth(2, 200); // 関数名
//     errorSheet.setColumnWidth(3, 100); // エラータイプ
//     errorSheet.setColumnWidth(4, 300); // エラーメッセージ
//     errorSheet.setColumnWidth(5, 400); // 詳細情報
    
//     // 元のシートに戻す
//     if (currentSheet) {
//       ss.setActiveSheet(currentSheet);
//     }
//   }
  
//   return errorSheet;
// }

// // エラーログを記録
// function logError_(functionName, errorType, errorMessage, details) {
//   try {
//     var errorSheet = getOrCreateErrorLogSheet_();
//     var lastRow = errorSheet.getLastRow();
//     var timestamp = new Date();
    
//     // エラー情報を配列にまとめる
//     var errorData = [
//       timestamp,
//       functionName || "不明",
//       errorType || "ERROR",
//       errorMessage || "エラーメッセージなし",
//       details || ""
//     ];
    
//     // エラーログシートに記録
//     errorSheet.getRange(lastRow + 1, 1, 1, errorData.length).setValues([errorData]);
    
//     // コンソールにも出力
//     console.error("エラーログ:", {
//       timestamp: timestamp,
//       functionName: functionName,
//       errorType: errorType,
//       errorMessage: errorMessage,
//       details: details
//     });
    
//   } catch (logError) {
//     // ログ記録自体が失敗した場合はコンソールに出力
//     console.error("エラーログの記録に失敗:", logError);
//   }
// }

// // 情報ログを記録（デバッグ用）
// function logInfo_(functionName, message, details) {
//   logError_(functionName, "INFO", message, details);
// }

// // 警告ログを記録
// function logWarning_(functionName, message, details) {
//   logError_(functionName, "WARNING", message, details);
// }

// // B110: オークファンの検索URL生成
// function generateAucfanUrl_(sheet, genre, keyword, period) {
//   try {
//     var encodedKeyword = encodeURIComponent(keyword);
    
//     // 正しいオークファンProのURL形式
//     var baseUrl = "https://pro.aucfan.com/search/list";
//     var params = [];
    
//     // 必須パラメータ
//     params.push("q=" + encodedKeyword);
//     params.push("mode=past_2y"); // 過去2年のデータ
//     params.push("site=y"); // ヤフオク
//     params.push("o=t1"); // ソート順（落札日の新しい順）
//     params.push("v=30"); // 表示件数
//     params.push("disp=list"); // リスト表示
//     params.push("disp_num=30"); // 表示件数
    
//     // 期間指定（G8）
//     if (period) {
//       var year, month, day;
      
//       // Dateオブジェクトの場合
//       if (period instanceof Date) {
//         year = period.getFullYear();
//         month = period.getMonth() + 1; // 0-indexed なので +1
//         day = period.getDate();
//         console.log("Dateオブジェクトから日付を取得: " + year + "/" + month + "/" + day);
//       } else {
//         // 文字列の日付形式を判定
//         var datePattern = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/;
//         var match = String(period).match(datePattern);
        
//         if (match) {
//           // 日付形式の場合（YYYY/MM/DD または YYYY-MM-DD）
//           year = parseInt(match[1]);
//           month = parseInt(match[2]);
//           day = parseInt(match[3]);
//         }
//       }
      
//       if (year && month && day) {
        
//         // 終了日
//         params.push("term_end_year=" + year);
//         params.push("term_end_month=" + month);
//         params.push("term_end=" + day);
        
//         // 開始日（1ヶ月前をデフォルトに）
//         var startDateObj = new Date(year, month - 1, day);
//         startDateObj.setMonth(startDateObj.getMonth() - 1);
        
//         params.push("term_begin_year=" + startDateObj.getFullYear());
//         params.push("term_begin_month=" + (startDateObj.getMonth() + 1));
//         params.push("term_begin=" + startDateObj.getDate());
//         params.push("term_month=1"); // 期間タイプ（1=カスタム期間）
        
//         console.log("期間指定: " + startDateObj.getFullYear() + "/" + (startDateObj.getMonth() + 1) + "/" + startDateObj.getDate() + 
//                     " 〜 " + year + "/" + month + "/" + day);
        
//       } else if (typeof period === 'string' && period.trim() !== '') {
//         // 期間文字列のチェック
//         var periodStr = period.trim();
//         if (periodStr === "1週間" || periodStr === "1week") {
//           params.push("period=1w");
//           console.log("期間指定: 1週間");
//         } else if (periodStr === "1ヶ月" || periodStr === "1month") {
//           params.push("period=1m");
//           console.log("期間指定: 1ヶ月");
//         } else if (periodStr === "3ヶ月" || periodStr === "3months") {
//           params.push("period=3m");
//           console.log("期間指定: 3ヶ月");
//         } else if (periodStr === "6ヶ月" || periodStr === "6months") {
//           params.push("period=6m");
//           console.log("期間指定: 6ヶ月");
//         } else if (periodStr === "1年" || periodStr === "1year") {
//           params.push("period=1y");
//           console.log("期間指定: 1年");
//         } else {
//           console.log("期間指定なし: デフォルト（過去2年）を使用");
//         }
//       } else {
//         // 期間文字列の場合はデフォルト（過去2年）を使用
//         console.log("期間指定なし: デフォルト（過去2年）を使用");
//       }
//     }
    
//     // その他のパラメータ
//     params.push("cid=0"); // カテゴリID（0=すべて）
//     params.push("page=1"); // ページ番号
    
//     // URL生成
//     var aucfanUrl = baseUrl + "?" + params.join("&");
    
//     sheet.getRange(110, COL_B).setValue(aucfanUrl); // B110に設定
//     console.log("B110にオークファンの検索URLを設定しました:", aucfanUrl);
//     logInfo_("generateAucfanUrl_", "オークファンURL生成", "キーワード: " + keyword + ", 期間: " + period);
    
//   } catch (e) {
//     console.error("オークファンURL生成エラー:", e.message);
//     logError_("generateAucfanUrl_", "ERROR", "オークファンURL生成失敗", e.message);
//     // エラーでも最低限のURLは設定
//     var fallbackUrl = "https://pro.aucfan.com/search/list?q=" + encodeURIComponent(keyword || "") + "&mode=past_2y&site=y";
//     sheet.getRange(110, COL_B).setValue(fallbackUrl);
//   }
// }

// // オークファンのカテゴリIDを取得
// function getAucfanCategoryId_(genre) {
//   // ジャンルマッピング（将来的に卸先シートから取得）
//   var categoryMap = {
//     "服-メンズ-ブランド": "2084027636",
//     "服-レディース-ブランド": "2084005354",
//     "バッグ-メンズ-ブランド": "2084024204",
//     "バッグ-レディース-ブランド": "2084062720",
//     "財布・小物-メンズ-ブランド": "2084292126",
//     "財布・小物-レディース-ブランド": "2084062696",
//     "時計-メンズ-ブランド": "2084053759",
//     "時計-レディース-ブランド": "2084053758",
//     "アクセサリー-メンズ-ブランド": "2084050107",
//     "アクセサリー-レディース-ブランド": "2084005422",
//     "靴-メンズ-ブランド": "2084055337",
//     "靴-レディース-ブランド": "2084047831",
//     "家電": "2084239269",
//     "カメラ": "2084261644",
//     "スマホ・タブレット": "2084039759",
//     "PC・周辺機器": "2084039785",
//     "ゲーム": "2084315442"
//   };
  
//   return categoryMap[genre] || "";
// }

// // オークファンの期間パラメータをフォーマット
// function formatAucfanPeriod_(period) {
//   var params = [];
  
//   if (period instanceof Date) {
//     // 日付オブジェクトの場合
//     var endDate = Utilities.formatDate(period, "JST", "yyyy-MM-dd");
//     var startDate = new Date(period);
//     startDate.setMonth(startDate.getMonth() - 3); // 3ヶ月前
//     var startDateStr = Utilities.formatDate(startDate, "JST", "yyyy-MM-dd");
    
//     params.push("date_from=" + startDateStr);
//     params.push("date_to=" + endDate);
//   } else if (typeof period === "string" && period.match(/^\d{4}-\d{2}-\d{2}$/)) {
//     // YYYY-MM-DD形式の文字列の場合
//     var endDate = period;
//     var parts = period.split("-");
//     var startDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
//     startDate.setMonth(startDate.getMonth() - 3);
//     var startDateStr = Utilities.formatDate(startDate, "JST", "yyyy-MM-dd");
    
//     params.push("date_from=" + startDateStr);
//     params.push("date_to=" + endDate);
//   } else if (period === "1週間" || period === "1week") {
//     params.push("period=1w");
//   } else if (period === "1ヶ月" || period === "1month") {
//     params.push("period=1m");
//   } else if (period === "3ヶ月" || period === "3months") {
//     params.push("period=3m");
//   } else if (period === "6ヶ月" || period === "6months") {
//     params.push("period=6m");
//   } else if (period === "1年" || period === "1year") {
//     params.push("period=1y");
//   }
  
//   return params;
// }

// // オークファンの高度な検索URL生成（価格範囲や条件を指定）
// function generateAucfanAdvancedUrl(keyword, options) {
//   options = options || {};
  
//   var baseUrl = "https://ssl.aucfan.com/aucfanpro/search";
//   var params = [];
  
//   // キーワード（必須）
//   if (!keyword) {
//     throw new Error("検索キーワードは必須です");
//   }
//   params.push("q=" + encodeURIComponent(keyword));
  
//   // カテゴリID
//   if (options.category) {
//     params.push("category=" + options.category);
//   }
  
//   // 価格範囲
//   if (options.minPrice) {
//     params.push("min=" + options.minPrice);
//   }
//   if (options.maxPrice) {
//     params.push("max=" + options.maxPrice);
//   }
  
//   // ソート順
//   var sortOrder = options.sort || "price_desc";
//   params.push("order=" + sortOrder);
//   // ソートオプション:
//   // - price_desc: 価格の高い順
//   // - price_asc: 価格の安い順
//   // - date_desc: 日付の新しい順
//   // - bid_desc: 入札数の多い順
  
//   // 商品の状態
//   if (options.condition) {
//     params.push("condition=" + options.condition);
//     // 1: 新品
//     // 2: 未使用
//     // 3: 中古
//   }
  
//   // サイト指定
//   if (options.site) {
//     params.push("site=" + options.site);
//     // yahoo: ヤフオク
//     // rakuten: 楽天
//     // mercari: メルカリ
//     // all: すべて
//   }
  
//   // 期間
//   if (options.period) {
//     var dateParams = formatAucfanPeriod_(options.period);
//     if (dateParams) {
//       params = params.concat(dateParams);
//     }
//   } else if (options.dateFrom && options.dateTo) {
//     params.push("date_from=" + options.dateFrom);
//     params.push("date_to=" + options.dateTo);
//   }
  
//   // オプション
//   if (options.closed !== false) {
//     params.push("closed=1"); // 落札済みのみ
//   }
  
//   // 表示件数
//   params.push("limit=" + (options.limit || 100));
  
//   // 完成したURLを返す
//   return baseUrl + "?" + params.join("&");
// }

// // テスト用：オークファンの検索URL生成例
// function testAucfanUrlGeneration() {
//   console.log("=== オークファン検索URL生成テスト ===");
  
//   // 例1: シンプルな検索
//   var url1 = generateAucfanAdvancedUrl("iPhone 15");
//   console.log("シンプル検索:", url1);
  
//   // 例2: 価格範囲指定
//   var url2 = generateAucfanAdvancedUrl("プラダ バッグ", {
//     minPrice: 50000,
//     maxPrice: 200000,
//     sort: "price_desc"
//   });
//   console.log("価格範囲指定:", url2);
  
//   // 例3: カテゴリと期間指定
//   var url3 = generateAucfanAdvancedUrl("ルイヴィトン", {
//     category: "2084062720", // バッグ-レディース-ブランド
//     period: "3ヶ月",
//     condition: "2" // 未使用
//   });
//   console.log("カテゴリ・期間指定:", url3);
  
//   // 例4: 複雑な条件
//   var url4 = generateAucfanAdvancedUrl("エルメス バーキン", {
//     minPrice: 1000000,
//     maxPrice: 5000000,
//     category: "2084062720",
//     site: "yahoo",
//     sort: "date_desc",
//     limit: 50,
//     dateFrom: "2024-01-01",
//     dateTo: "2024-12-31"
//   });
//   console.log("複雑な条件:", url4);
  
//   // スプレッドシートに例を出力
//   var sheet = SpreadsheetApp.getActiveSheet();
//   sheet.getRange("G10").setValue("=== オークファン検索URL例 ===");
//   sheet.getRange("G11").setValue(url1);
//   sheet.getRange("G12").setValue(url2);
//   sheet.getRange("G13").setValue(url3);
//   sheet.getRange("G14").setValue(url4);
  
//   console.log("\nG10〜G14に検索URL例を出力しました");
// }

// // オークファンURL生成のテスト関数（B110確認用）
// function testAucfanB110() {
//   console.log("=== オークファンB110テスト開始 ===");
  
//   try {
//     var sheet = SpreadsheetApp.getActiveSheet();
    
//     // G7とG8の値を取得（タイムアウト対策）
//     var keyword = "";
//     var period = "";
    
//     try {
//       keyword = sheet.getRange("G7").getValue();
//       period = sheet.getRange("G8").getValue();
//     } catch (e) {
//       console.error("セルからの値取得エラー:", e.message);
//       SpreadsheetApp.getUi().alert("エラー: セルからの値取得に失敗しました\n" + e.message);
//       return;
//     }
    
//     console.log("G7（キーワード）:", keyword);
//     console.log("G8（期間）:", period);
    
//     // 入力チェック
//     if (!keyword || keyword.toString().trim() === "") {
//       console.warn("キーワードが入力されていません");
//       SpreadsheetApp.getUi().alert("警告: G7にキーワードを入力してください");
//       return;
//     }
    
//     // URL生成（タイムアウト対策）
//     try {
//       generateAucfanUrl_(sheet, "", keyword, period);
//     } catch (e) {
//       console.error("URL生成エラー:", e.message);
//       SpreadsheetApp.getUi().alert("エラー: URL生成に失敗しました\n" + e.message);
//       return;
//     }
    
//     // B110の値を確認（タイムアウト対策）
//     var generatedUrl = "";
//     try {
//       generatedUrl = sheet.getRange(110, 2).getValue();
//     } catch (e) {
//       console.error("B110の値取得エラー:", e.message);
//       SpreadsheetApp.getUi().alert("エラー: B110の値取得に失敗しました\n" + e.message);
//       return;
//     }
    
//     console.log("B110に生成されたURL:", generatedUrl);
    
//     if (generatedUrl) {
//       console.log("✅ B110にURLが生成されました!");
//       console.log("URL:", generatedUrl);
      
//       // オプション：生成されたURLの妥当性チェック
//       if (generatedUrl.indexOf("pro.aucfan.com") === -1) {
//         console.warn("生成されたURLがオークファンのURLではない可能性があります");
//       }
      
//       // アラートの代わりにトーストメッセージを使用（非ブロッキング）
//       try {
//         SpreadsheetApp.getActiveSpreadsheet().toast("✅ B110にURLが生成されました!", "成功", 3);
//       } catch (toastError) {
//         console.log("トースト表示エラー:", toastError.message);
//       }
//     } else {
//       console.error("❌ B110にURLが生成されませんでした。");
//       console.log("以下を確認してください:");
//       console.log("・G7にキーワードが入力されているか");
//       console.log("・G8の期間形式が正しいか");
      
//       // アラートの代わりにトーストメッセージを使用（非ブロッキング）
//       try {
//         SpreadsheetApp.getActiveSpreadsheet().toast("❌ B110にURLが生成されませんでした", "エラー", 5);
//       } catch (toastError) {
//         console.log("トースト表示エラー:", toastError.message);
//       }
//     }
    
//   } catch (e) {
//     console.error("testAucfanB110エラー:", e.message);
//     console.error("スタックトレース:", e.stack);
//     SpreadsheetApp.getUi().alert("予期しないエラーが発生しました:\n\n" + e.message + "\n\n詳細はログを確認してください");
//   } finally {
//     console.log("=== オークファンB110テスト終了 ===");
//   }
// }

// // よく使うオークファン検索URLの簡易生成関数
// function quickAucfanUrl(keyword) {
//   // 最も基本的な検索URL（過去2年、ヤフオク、落札日新しい順）
//   return "https://pro.aucfan.com/search/list?q=" + 
//          encodeURIComponent(keyword) + 
//          "&mode=past_2y&site=y&o=t1&v=30&disp=list";
// }

// // ブランド品検索用URL生成
// function aucfanBrandUrl(brandName, itemType, options) {
//   options = options || {};
  
//   var keyword = brandName;
//   if (itemType) {
//     keyword += " " + itemType;
//   }
  
//   // ブランド品検索のデフォルト設定
//   var defaultOptions = {
//     minPrice: 10000, // 最低価格1万円
//     sort: "price_desc",
//     closed: true,
//     limit: 100
//   };
  
//   // オプションをマージ
//   for (var key in defaultOptions) {
//     if (!(key in options)) {
//       options[key] = defaultOptions[key];
//     }
//   }
  
//   return generateAucfanAdvancedUrl(keyword, options);
// }

// // スプレッドシートのセルからパラメータを読み取ってURL生成
// function generateAucfanFromSheet() {
//   var sheet = SpreadsheetApp.getActiveSheet();
  
//   // G6: ジャンル、G7: キーワード、G8: 期間
//   var genre = sheet.getRange("G6").getValue();
//   var keyword = sheet.getRange("G7").getValue();
//   var period = sheet.getRange("G8").getValue();
  
//   // オプションパラメータ（必要に応じて追加）
//   var minPrice = sheet.getRange("G9").getValue() || "";
//   var maxPrice = sheet.getRange("G10").getValue() || "";
  
//   if (!keyword) {
//     SpreadsheetApp.getUi().alert("G7に検索キーワードを入力してください");
//     return;
//   }
  
//   var options = {
//     category: getAucfanCategoryId_(genre),
//     period: period
//   };
  
//   if (minPrice) options.minPrice = minPrice;
//   if (maxPrice) options.maxPrice = maxPrice;
  
//   var url = generateAucfanAdvancedUrl(keyword, options);
  
//   // B110に設定
//   sheet.getRange(110, COL_B).setValue(url);
  
//   SpreadsheetApp.getUi().alert("オークファンのURLを生成しました:\n" + url);
  
//   return url;
// }

// // エラーログシートのテスト関数
// function testErrorLog() {
//   logInfo_("testErrorLog", "テスト開始", "エラーログシートの動作確認");
//   logWarning_("testErrorLog", "警告テスト", "これは警告のテストです");
//   logError_("testErrorLog", "ERROR", "エラーテスト", "これはエラーのテストです");
//   logInfo_("testErrorLog", "テスト完了", "エラーログシートを確認してください");
// }

// // URLの配置状況を診断する関数
// function diagnoseSiteUrls() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getActiveSheet();

//   console.log("=== サイトURL診断 ===");

//   // B11: SBA
//   var sbaUrl = sheet.getRange(11, COL_B).getValue();
//   console.log("B11 (SBA):", sbaUrl || "空");

//   // B44: エコリング
//   var ecoringUrl = sheet.getRange(44, COL_B).getValue();
//   console.log("B44 (エコリング):", ecoringUrl || "空");

//   // B77: 楽天
//   var rakutenUrl = sheet.getRange(77, COL_B).getValue();
//   console.log("B77 (楽天):", rakutenUrl || "空");

//   // B110: オークファン
//   var aucfanUrl = sheet.getRange(110, COL_B).getValue();
//   console.log("B110 (オークファン):", aucfanUrl || "空");

//   // B143: ヤフオク
//   var yahooUrl = sheet.getRange(143, COL_B).getValue();
//   console.log("B143 (ヤフオク):", yahooUrl || "空");

//   // エラーログにも記録
//   logInfo_("diagnoseSiteUrls", "URL診断実行",
//     "SBA: " + (sbaUrl ? "有" : "無") +
//     ", エコリング: " + (ecoringUrl ? "有" : "無") +
//     ", 楽天: " + (rakutenUrl ? "有" : "無") +
//     ", オークファン: " + (aucfanUrl ? "有" : "無") +
//     ", ヤフオク: " + (yahooUrl ? "有" : "無"));
// }

// // ヤフオクの動作をテストする関数
// function testYahooAuction() {
//   try {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var sheet = ss.getActiveSheet();
    
//     logInfo_("testYahooAuction", "ヤフオクテスト開始", "");
    
//     // テスト用のURL（CELINEで検索）
//     var testUrl = "https://auctions.yahoo.co.jp/closedsearch/closedsearch?auccat=23000&tab_ex=commerce&ei=utf-8&aq=-1&oq=&sc_i=&p=CELINE";
    
//     console.log("=== ヤフオクテスト開始 ===");
//     console.log("テストURL:", testUrl);
    
//     // HTMLを取得
//     console.log("1. HTMLの取得を開始...");
//     var html = fetchYahooAuctionHtml_(testUrl);
//     console.log("HTMLの長さ:", html.length);
//     console.log("HTMLに'Product'クラスが含まれているか:", html.indexOf('class="Product"') > -1);
//     console.log("HTMLに'Products__items'が含まれているか:", html.indexOf('Products__items') > -1);
    
//     // HTMLの一部を表示（商品部分）
//     var productStart = html.indexOf('<li class="Product"');
//     if (productStart > -1) {
//       console.log("最初の商品HTMLの一部:");
//       console.log(html.substring(productStart, productStart + 1000));
//     }
    
//     // パース
//     console.log("\n2. HTMLのパースを開始...");
//     var items = parseYahooAuctionFromHtml_(html);
//     console.log("パース結果: " + items.length + "件の商品");
    
//     // 最初の3商品を詳細表示
//     if (items.length > 0) {
//       console.log("\n=== パースした商品データ ===");
//       for (var i = 0; i < Math.min(3, items.length); i++) {
//         console.log("\n商品 " + (i + 1) + ":");
//         console.log("  タイトル:", items[i].title);
//         console.log("  URL:", items[i].detailUrl);
//         console.log("  画像:", items[i].imageUrl);
//         console.log("  価格:", items[i].price);
//         console.log("  日付:", items[i].date);
//         console.log("  ショップ:", items[i].shop);
//       }
//     } else {
//       console.log("商品が見つかりませんでした。パーサーに問題がある可能性があります。");
//     }
    
//     logInfo_("testYahooAuction", "ヤフオクテスト完了", items.length + "件の商品を取得");
    
//     // B143にテストURLを設定（オプション）
//     // sheet.getRange(143, COL_B).setValue(testUrl);
    
//     return items;
    
//   } catch (e) {
//     console.error("ヤフオクテストエラー:", e.message);
//     logError_("testYahooAuction", "ERROR", "ヤフオクテストでエラー発生", e.message + "\n" + e.stack);
//     throw e;
//   }
// }

// // エコリングの動作をテストする関数
// function testEcoRing() {
//   try {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var sheet = ss.getActiveSheet();
    
//     logInfo_("testEcoRing", "エコリングテスト開始", "");
    
//     // テスト用のURL（ヴィトンで検索）
//     var testUrl = "https://www.ecoauc.com/client/market-prices?limit=50&sortKey=1&q=%E3%83%B4%E3%82%A3%E3%83%88%E3%83%B3&low=&high=&master_item_brands=&master_item_categories%5B0%5D=8&master_item_shapes=&target_start_year=2025&target_start_month=1&target_end_year=2025&target_end_month=8&master_invoice_setting_id=0&method=1&master_item_ranks=&tab=general&tableType=grid";
    
//     console.log("=== エコリングテスト開始 ===");
//     console.log("テストURL:", testUrl);
    
//     // HTMLを取得
//     console.log("1. HTMLの取得を開始...");
//     var html = fetchEcoringHtml_(testUrl);
//     console.log("HTMLの長さ:", html.length);
//     console.log("HTMLに'show-case-title-block'が含まれているか:", html.indexOf('show-case-title-block') > -1);
//     console.log("HTMLに'show-value'が含まれているか:", html.indexOf('show-value') > -1);
//     console.log("HTMLに'col-sm-6 col-md-4 col-lg-3'が含まれているか:", html.indexOf('col-sm-6 col-md-4 col-lg-3') > -1);
    
//     // HTMLの一部を表示（商品部分）
//     var productStart = html.indexOf('<div class="col-sm-6 col-md-4 col-lg-3">');
//     if (productStart > -1) {
//       console.log("最初の商品HTMLの一部:");
//       console.log(html.substring(productStart, productStart + 1000));
//     }
    
//     // パース
//     console.log("\n2. HTMLのパースを開始...");
//     var items = parseEcoringFromHtml_(html);
//     console.log("パース結果: " + items.length + "件の商品");
    
//     // 最初の3商品を詳細表示
//     if (items.length > 0) {
//       console.log("\n=== パースした商品データ ===");
//       for (var i = 0; i < Math.min(3, items.length); i++) {
//         console.log("\n商品 " + (i + 1) + ":");
//         console.log("  タイトル:", items[i].title);
//         console.log("  URL:", items[i].detailUrl);
//         console.log("  画像:", items[i].imageUrl);
//         console.log("  価格:", items[i].price);
//         console.log("  日付:", items[i].date);
//         console.log("  ランク:", items[i].rank);
//         console.log("  ショップ:", items[i].shop);
//       }
//     } else {
//       console.log("商品が見つかりませんでした。パーサーまたはログインに問題がある可能性があります。");
//     }
    
//     logInfo_("testEcoRing", "エコリングテスト完了", items.length + "件の商品を取得");
    
//     // B44にテストURLを設定（オプション）
//     // sheet.getRange(44, COL_B).setValue(testUrl);
    
//     return items;
    
//   } catch (e) {
//     console.error("エコリングテストエラー:", e.message);
//     logError_("testEcoRing", "ERROR", "エコリングテストでエラー発生", e.message + "\n" + e.stack);
//     throw e;
//   }
// }

// // カテゴリID取得のテスト関数
// function testCategoryIds() {
//   var genre = "レディースバッグ"; // テスト用のジャンル（実際のジャンルに変更してください）
  
//   console.log("=== カテゴリID取得テスト ===");
//   console.log("テストジャンル:", genre);
  
//   try {
//     var sbaId = getSbaCategoryId_(genre);
//     console.log("SBAカテゴリID:", sbaId);
//   } catch (e) {
//     console.error("SBAカテゴリID取得エラー:", e.message);
//   }
  
//   try {
//     var rakutenId = getRakutenCategoryId_(genre);
//     console.log("楽天カテゴリID:", rakutenId);
//   } catch (e) {
//     console.error("楽天カテゴリID取得エラー:", e.message);
//   }
  
//   try {
//     var ecoringIds = getEcoringCategoryIds_(genre);
//     console.log("エコリングカテゴリID:", ecoringIds);
//   } catch (e) {
//     console.error("エコリングカテゴリID取得エラー:", e.message);
//   }
  
//   try {
//     var yahooId = getYahooCategoryId_(genre);
//     console.log("ヤフオクカテゴリID:", yahooId);
//   } catch (e) {
//     console.error("ヤフオクカテゴリID取得エラー:", e.message);
//   }
// }

// // オークファンの価格取得テスト関数（固定URLを使用）
// function testAucfanPriceExtraction() {
//   console.log("╔══════════════════════════════════════════════════════╗");
//   console.log("║     🌐 オークファン ログイン状態確認テスト           ║");
//   console.log("╚══════════════════════════════════════════════════════╝");
//   console.log("\n⏳ テストを開始します...");
  
//   var startTime = new Date().getTime();
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
  
//   console.log("=== オークファン価格取得テスト（固定URL）===");
  
//   try {
//     // 固定のテストURL（グッチのバッグを検索）
//     var testUrl = "https://pro.aucfan.com/search/list?q=%E3%82%B0%E3%83%83%E3%83%81%20%E3%83%90%E3%83%83%E3%82%B0&mode=past_2y&site=y&o=t1&v=30&disp=list&disp_num=30&cid=0&page=1";
//     console.log("テストURL:", testUrl);
    
//     // Cookiesシートからクッキーを読み込む
//     var cookies = loadCookiesFromSheet();
//     console.log("クッキー読み込み完了:", cookies.length + "個");
    
//     if (cookies.length === 0) {
//       console.error("❌ クッキーが見つかりません");
//       console.error("=== 対処法 ===");
//       console.error("1. ターミナルを開く");
//       console.error("2. cd /Users/kazuyukijimbo/aicon");
//       console.error("3. python aucfan_sheet_writer.py");
//       console.error("4. ログインとreCAPTCHAを完了");
//       console.error("5. クッキーがスプレッドシートに保存されたら再実行");
      
//       SpreadsheetApp.getActiveSpreadsheet().toast(
//         "クッキーがありません。\n\nターミナルで以下を実行:\npython aucfan_sheet_writer.py", 
//         "❌ 要ログイン", 
//         15
//       );
//       return;
//     }
    
//     // URLにアクセスしてHTMLを取得
//     var response = UrlFetchApp.fetch(testUrl, {
//       muteHttpExceptions: true,
//       headers: {
//         "Cookie": formatCookieHeader(cookies),
//         "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
//         "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8"
//       },
//       timeout: 30 // 30秒のタイムアウト設定
//     });
    
//     var statusCode = response.getResponseCode();
//     console.log("レスポンスステータス:", statusCode);
    
//     if (statusCode === 200) {
//       var html = response.getContentText();
//       console.log("HTMLサイズ:", html.length + "文字");
      
//       // ログインページにリダイレクトされていないかチェック
//       if (html.indexOf("ログイン") > -1 && html.indexOf("パスワード") > -1) {
//         console.error("ログインページが表示されました。クッキーが無効またはセッションが期限切れです。");
//         console.log("HTMLの一部（ログイン検出）:", html.substring(0, 500));
        
//         // クッキーの詳細情報を表示
//         console.log("=== クッキー情報 ===");
//         console.log("クッキー数:", cookies.length);
//         var sessionCookie = cookies.find(function(c) { return c.name === 'PHPSESSID' || c.name === 'session'; });
//         if (sessionCookie) {
//           console.log("セッションクッキー:", sessionCookie.name, "値の長さ:", sessionCookie.value.length);
//           if (sessionCookie.expires) {
//             var expiryDate = new Date(sessionCookie.expires * 1000);
//             console.log("有効期限:", expiryDate.toLocaleString());
//             if (expiryDate < new Date()) {
//               console.error("❌ セッションクッキーの有効期限が切れています！");
//             }
//           }
//         } else {
//           console.error("❌ セッションクッキーが見つかりません");
//         }
        
//         // alertの代わりにtoastとエラーメッセージ
//         try {
//           SpreadsheetApp.getActiveSpreadsheet().toast(
//             "セッション切れです。aucfan_sheet_writer.pyで再ログインしてください。", 
//             "❌ 認証エラー", 
//             10
//           );
//         } catch (e) {
//           console.log("Toast表示エラー:", e.message);
//         }
        
//         console.error("\n=== 対処法 ===");
//         console.error("1. ターミナルで以下を実行:");
//         console.error("   cd /Users/kazuyukijimbo/aicon");
//         console.error("   python aucfan_sheet_writer.py");
//         console.error("2. ログイン完了後、再度このテストを実行してください");
        
//         return;
//       }
      
//       // パース実行
//       var items = parseAucfanFromHtml_(html);
//       console.log("取得商品数:", items.length);
      
//       // 最初の5商品の価格情報を表示
//       console.log("\n=== 価格取得結果 ===");
//       for (var i = 0; i < Math.min(5, items.length); i++) {
//         console.log("\n商品" + (i + 1) + ":");
//         console.log("  タイトル:", items[i].title ? items[i].title.substring(0, 50) + "..." : "タイトルなし");
//         console.log("  価格:", items[i].price || "価格なし");
//         console.log("  URL:", items[i].detailUrl || "URLなし");
//         console.log("  落札日:", items[i].date || "日付なし");
//       }
      
//       // 価格がない商品をカウント
//       var noPriceCount = 0;
//       items.forEach(function(item) {
//         if (!item.price || item.price === "0") {
//           noPriceCount++;
//         }
//       });
      
//       if (noPriceCount > 0) {
//         console.log("\n価格が取得できなかった商品数:", noPriceCount + "/" + items.length);
//       } else {
//         console.log("\n全商品で価格を取得できました");
//       }
      
//       // 結果をAucfanResultsシートに保存
//       if (items.length > 0) {
//         saveAucfanResultsToSheet(items, "グッチ バッグ（テスト）");
//         console.log("\n結果をAucfanResultsシートに保存しました");
//       }
//     } else {
//       console.error("リクエストに失敗しました。ステータスコード:", statusCode);
//       var errorText = response.getContentText();
//       console.log("エラーレスポンス:", errorText.substring(0, 500) + "...");
//     }
    
//   } catch (e) {
//     console.error("テストエラー:", e.message);
//   }
// }

// // 簡単なテスト実行用のエイリアス関数
// function testAucfan() {
//   testAucfanPriceExtraction();
// }

// // 総合的なログイン状態とクッキー診断を実行する関数
// function checkLoginStatus() {
//   console.log("╔══════════════════════════════════════════════════════╗");
//   console.log("║       🔐 総合ログイン状態チェックを開始              ║");
//   console.log("╚══════════════════════════════════════════════════════╝");
//   console.log("");
  
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var allChecksPassed = true;
//   var statusMessages = [];
  
//   // Step 1: クッキー診断
//   console.log("【ステップ 1/2】クッキー診断");
//   console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  
//   var cookies = loadCookiesFromSheet();
  
//   if (cookies.length === 0) {
//     allChecksPassed = false;
//     statusMessages.push("❌ クッキーなし");
//   } else {
//     var validation = validateCookies(cookies);
//     console.log(validation.message);
    
//     if (validation.detailedMessage.length > 0) {
//       validation.detailedMessage.forEach(function(msg) {
//         console.log("  → " + msg);
//       });
//     }
    
//     if (!validation.isValid) {
//       allChecksPassed = false;
//       statusMessages.push(validation.message);
//     } else {
//       statusMessages.push("✅ クッキー正常");
//     }
//   }
  
//   console.log("");
  
//   // Step 2: 実際のログイン状態テスト（クッキーがある場合のみ）
//   if (cookies.length > 0) {
//     console.log("【ステップ 2/2】実際のログイン状態確認");
//     console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
//     console.log("📡 オークファンサーバーにアクセス中...");
    
//     try {
//       var testUrl = "https://pro.aucfan.com/";
//       var response = UrlFetchApp.fetch(testUrl, {
//         muteHttpExceptions: true,
//         headers: {
//           "Cookie": formatCookieHeader(cookies),
//           "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
//         },
//         timeout: 10
//       });
      
//       var statusCode = response.getResponseCode();
//       var html = response.getContentText();
      
//       if (statusCode === 200 && html.indexOf("ログイン") > -1 && html.indexOf("パスワード") > -1) {
//         console.log("❌ ログインページが表示されました（未認証）");
//         allChecksPassed = false;
//         statusMessages.push("❌ ログイン必要");
//       } else if (statusCode === 200) {
//         console.log("✅ ログイン状態確認済み（認証済み）");
//         statusMessages.push("✅ ログイン済み");
//       } else {
//         console.log("⚠️ サーバーエラー（ステータス: " + statusCode + "）");
//         allChecksPassed = false;
//         statusMessages.push("⚠️ サーバーエラー");
//       }
//     } catch (e) {
//       console.error("❌ 接続エラー: " + e.message);
//       allChecksPassed = false;
//       statusMessages.push("❌ 接続エラー");
//     }
//   } else {
//     console.log("【ステップ 2/2】スキップ（クッキーがないため）");
//   }
  
//   console.log("");
//   console.log("╔══════════════════════════════════════════════════════╗");
  
//   if (allChecksPassed) {
//     console.log("║         ✅ 全てのチェックに合格しました              ║");
//     console.log("╚══════════════════════════════════════════════════════╝");
//     console.log("\n🎉 オークファンAPIを使用する準備が整いました！");
    
//     ss.toast(
//       "✅ ログイン状態: 正常\n\n" +
//       statusMessages.join("\n"),
//       "🟢 準備完了",
//       8
//     );
//   } else {
//     console.log("║         ⚠️ 一部のチェックに失敗しました             ║");
//     console.log("╚══════════════════════════════════════════════════════╝");
//     console.log("\n📝 対処法:");
//     console.log("  1. ターミナルで以下を実行:");
//     console.log("     cd /Users/kazuyukijimbo/aicon");
//     console.log("     python aucfan_sheet_writer.py");
//     console.log("  2. ログインを完了してください");
    
//     ss.toast(
//       "⚠️ ログイン状態に問題があります\n\n" +
//       statusMessages.join("\n") + "\n\n" +
//       "aucfan_sheet_writer.pyで再ログインしてください",
//       "🔴 要対応",
//       15
//     );
//   }
  
//   // エラーログに記録
//   logInfo_("checkLoginStatus", "ログイン状態チェック実行", statusMessages.join(", "));
  
//   return allChecksPassed;
// }

// // クッキー診断関数（改善版）
// function diagnoseCookies() {
//   console.log("╔══════════════════════════════════════════╗");
//   console.log("║       🔍 クッキー診断を開始します        ║");
//   console.log("╚══════════════════════════════════════════╝");
  
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var cookies = loadCookiesFromSheet();
//   var statusMessages = [];
  
//   if (cookies.length === 0) {
//     console.error("\n❌ エラー: クッキーが保存されていません");
//     console.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
//     console.error("📝 対処法:");
//     console.error("  1. ターミナルを開く");
//     console.error("  2. 以下のコマンドを実行:");
//     console.error("     cd /Users/kazuyukijimbo/aicon");
//     console.error("     python aucfan_sheet_writer.py");
//     console.error("  3. ブラウザでログインを完了");
//     console.error("  4. reCAPTCHAを解決（表示された場合）");
//     console.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    
//     ss.toast(
//       "クッキーが見つかりません。\n\n" +
//       "ターミナルで以下を実行してください:\n" +
//       "python aucfan_sheet_writer.py",
//       "❌ クッキー未保存",
//       15
//     );
//     return;
//   }
  
//   console.log("\n✅ クッキー検出: " + cookies.length + "個のクッキーを発見");
//   statusMessages.push("クッキー数: " + cookies.length + "個");
  
//   // クッキーの有効性チェック
//   var validation = validateCookies(cookies);
  
//   console.log("\n📊 診断結果:");
//   console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  
//   if (validation.isValid) {
//     console.log("  ✅ ステータス: 正常");
//     console.log("  ✅ " + validation.message);
//     statusMessages.push("ステータス: 正常 ✅");
    
//     ss.toast(
//       "✅ クッキーは正常です\n\n" +
//       "クッキー数: " + cookies.length + "個\n" +
//       "セッション: 有効",
//       "🟢 診断成功",
//       8
//     );
//   } else {
//     console.error("  ❌ ステータス: 異常");
//     console.error("  ❌ " + validation.message);
//     statusMessages.push("ステータス: 異常 ❌");
    
//     // エラーの詳細を表示
//     if (validation.expiredCookies.length > 0) {
//       console.error("\n  ⚠️  期限切れクッキー:");
//       validation.expiredCookies.forEach(function(name) {
//         console.error("    - " + name);
//       });
//     }
    
//     if (!validation.hasSessionCookie) {
//       console.error("\n  ⚠️  セッションクッキーが見つかりません");
//       console.error("  📝 再ログインが必要です");
//     }
    
//     ss.toast(
//       "❌ クッキーに問題があります\n\n" +
//       validation.message + "\n\n" +
//       "python aucfan_sheet_writer.pyで再ログインしてください",
//       "🔴 要対応",
//       15
//     );
//   }
  
//   console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  
//   // 詳細情報を表示
//   console.log("\n📋 クッキー詳細情報:");
//   console.log("┌─────┬────────────────────────┬─────────────────┬──────────┐");
//   console.log("│ No. │ クッキー名             │ ドメイン        │ 状態     │");
//   console.log("├─────┼────────────────────────┼─────────────────┼──────────┤");
  
//   cookies.forEach(function(cookie, index) {
//     var status = "有効";
//     if (cookie.expiry) {
//       var expiryDate = new Date(cookie.expiry * 1000);
//       if (expiryDate < new Date()) {
//         status = "期限切れ";
//       }
//     }
    
//     var nameDisplay = (cookie.name + "                        ").substring(0, 22);
//     var domainDisplay = (cookie.domain + "                ").substring(0, 15);
//     var statusDisplay = status === "有効" ? "✅ 有効  " : "❌ " + status;
    
//     console.log("│ " + String(index + 1).padStart(3, " ") + " │ " + nameDisplay + " │ " + domainDisplay + " │ " + statusDisplay + " │");
    
//     if (cookie.expiry) {
//       var expiryDate = new Date(cookie.expiry * 1000);
//       console.log("│     │   有効期限: " + expiryDate.toLocaleString('ja-JP') + " │          │");
//     }
//   });
  
//   console.log("└─────┴────────────────────────┴─────────────────┴──────────┘");
  
//   console.log("\n╔══════════════════════════════════════════╗");
//   console.log("║       ✅ 診断が完了しました              ║");
//   console.log("╚══════════════════════════════════════════╝");
  
//   // エラーログにも記録
//   logInfo_("diagnoseCookies", "クッキー診断実行", statusMessages.join("、"));
// }

// // ヤフオクの診断機能
// function testYahooAuctionDiagnosis() {
//   console.log("=== ヤフオク診断開始 ===");
//   console.log("最大取得商品数: 50件");
//   console.log("最大表示商品数: " + MAX_OUTPUT_ITEMS + "件（C列〜V列）");
//   console.log("\n実際のWebページと表示が異なる理由:");
//   console.log("1. 商品名が取得できない商品はスキップされます");
//   console.log("2. 詳細URL・価格・日付がすべてない商品もスキップされます");
//   console.log("3. 最大表示数は20件に制限されています");
//   console.log("4. スキップされた商品により、順番がずれることがあります");
//   console.log("\n詳細は testYahooAuction() を実行してください");
  
//   // 実際にテストしてみる
//   try {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var sheet = ss.getActiveSheet();
//     var yahooUrl = sheet.getRange(143, COL_B).getValue();
    
//     if (yahooUrl && yahooUrl.indexOf("auctions.yahoo.co.jp") > -1) {
//       console.log("\n現在のヤフオクURL:", yahooUrl);
//       console.log("このURLでテストを実行しています...");
      
//       var html = fetchYahooAuctionHtml_(yahooUrl);
//       var items = parseYahooAuctionFromHtml_(html);
      
//       console.log("\n=== 診断結果 ===");
//       console.log("取得できた商品数:", items.length);
//       console.log("実際に表示される商品数:", Math.min(items.length, MAX_OUTPUT_ITEMS));
      
//       if (items.length > MAX_OUTPUT_ITEMS) {
//         console.log("\n注意: " + (items.length - MAX_OUTPUT_ITEMS) + "件の商品が表示上限により省略されます");
//       }
//     } else {
//       console.log("\nB143にヤフオクのURLが設定されていません");
//     }
//   } catch (e) {
//     console.error("診断エラー:", e.message);
//   }
// }

// // 全サイトの動作を簡易的にテストする関数
// function testAllSites() {
//   console.log("=== 全サイトテスト開始 ===");
//   logInfo_("testAllSites", "全サイトテスト開始", "");
  
//   var results = {
//     yahoo: { success: false, items: 0, error: null },
//     ecoring: { success: false, items: 0, error: null }
//   };
  
//   // ヤフオクのテスト
//   console.log("\n--- ヤフオクテスト ---");
//   try {
//     var yahooItems = testYahooAuction();
//     results.yahoo.success = true;
//     results.yahoo.items = yahooItems.length;
//   } catch (e) {
//     results.yahoo.error = e.message;
//     console.error("ヤフオクテストエラー:", e.message);
//   }
  
//   // エコリングのテスト
//   console.log("\n--- エコリングテスト ---");
//   try {
//     var ecoItems = testEcoRing();
//     results.ecoring.success = true;
//     results.ecoring.items = ecoItems.length;
//   } catch (e) {
//     results.ecoring.error = e.message;
//     console.error("エコリングテストエラー:", e.message);
//   }
  
//   // 結果サマリー
//   console.log("\n=== テスト結果サマリー ===");
//   console.log("ヤフオク:", results.yahoo.success ? "成功（" + results.yahoo.items + "件）" : "失敗（" + results.yahoo.error + "）");
//   console.log("エコリング:", results.ecoring.success ? "成功（" + results.ecoring.items + "件）" : "失敗（" + results.ecoring.error + "）");
  
//   logInfo_("testAllSites", "全サイトテスト完了", 
//     "ヤフオク: " + (results.yahoo.success ? results.yahoo.items + "件" : "失敗") + 
//     ", エコリング: " + (results.ecoring.success ? results.ecoring.items + "件" : "失敗"));
  
//   return results;
// }

// /** ============== 共通ユーティリティ ============== **/

// // ROW_HTML_INPUT 以降に貼られたHTMLを結合して取得（改行で連結）
// // 楽天URLの場合は自動的にフェッチ
// function readHtmlFromRow_(sheet) {
//   var last = sheet.getLastRow();
//   if (last < ROW_HTML_INPUT)
//     throw new Error("B" + ROW_HTML_INPUT + "にHTMLがありません");
//   var vals = sheet
//     .getRange(ROW_HTML_INPUT, COL_B, last - ROW_HTML_INPUT + 1, 1)
//     .getValues()
//     .map(function (r) {
//       return r[0];
//     });
//   var joined = vals
//     .filter(function (v) {
//       return v !== "" && v !== null;
//     })
//     .join("\n");
//   if (!joined)
//     throw new Error("B" + ROW_HTML_START + "から下にHTMLが見つかりません");

//   // URLかどうかチェック
//   var trimmed = joined.trim();
//   if (trimmed.match(/^https?:\/\/.*rakuten\.co\.jp\//i)) {
//     return fetchRakutenHtml_(trimmed);
//   }
//   if (trimmed.match(/^https?:\/\/.*starbuyers-global-auction\.com\//i)) {
//     return fetchStarBuyersHtml_(trimmed);
//   }
//   if (trimmed.match(/^https?:\/\/.*aucfan\.com\//i)) {
//     return fetchAucfanHtml_(trimmed);
//   }
//   if (trimmed.match(/^https?:\/\/.*ecoauc\.com\//i)) {
//     return fetchEcoringHtml_(trimmed);
//   }
//   if (trimmed.match(/^https?:\/\/.*auctions\.yahoo\.co\.jp\//i)) {
//     return fetchYahooAuctionHtml_(trimmed);
//   }

//   return String(joined);
// }

// // 全角数字→半角、カンマ/空白除去
// function normalizeNumber_(s) {
//   if (s == null) return "";
//   s = String(s);
//   s = s.replace(/[！-～]/g, function (ch) {
//     return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
//   });
//   s = s.replace(/\u3000/g, " ");
//   s = s.replace(/[¥￥円]/g, "");
//   s = s.replace(/[ ,\s]/g, "");
//   // 数字のみで構成されているかチェック
//   if (s && /^\d+$/.test(s)) {
//     return s;
//   }
//   return ""; // 数値でない場合は空文字列を返す
// }

// // 楽天URLからHTMLをフェッチ
// function fetchRakutenHtml_(url) {
//   try {
//     var options = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
//       },
//       muteHttpExceptions: true,
//     };

//     var response = UrlFetchApp.fetch(url, options);
//     var html = getResponseTextWithBestCharset_(response);

//     if (response.getResponseCode() !== 200) {
//       throw new Error(
//         "楽天からのHTMLの取得に失敗しました。ステータスコード: " +
//           response.getResponseCode()
//       );
//     }

//     return html;
//   } catch (e) {
//     throw new Error("楽天URLからのHTMLフェッチエラー: " + e.message);
//   }
// }

// // Star Buyers AuctionのURLからHTMLをフェッチ（ログイン必要）
// function fetchStarBuyersHtml_(url) {
//   try {
//     // ランダムな待機時間（1-3秒）
//     Utilities.sleep(1000 + Math.floor(Math.random() * 2000));

//     // ログイン情報（ベタ打ち）
//     var loginEmail = "inui.hur@gmail.com";
//     var loginPassword = "hur22721";

//     // ログインURL
//     var loginUrl = "https://www.starbuyers-global-auction.com/login";

//     // ステップ1: ログインページにアクセスしてCSRFトークンとクッキーを取得
//     var loginPageOptions = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//         "sec-ch-ua":
//           '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
//         "sec-ch-ua-mobile": "?0",
//         "sec-ch-ua-platform": '"macOS"',
//         "sec-fetch-dest": "document",
//         "sec-fetch-mode": "navigate",
//         "sec-fetch-site": "none",
//         "sec-fetch-user": "?1",
//         "upgrade-insecure-requests": "1",
//       },
//       muteHttpExceptions: true,
//       followRedirects: false,
//     };

//     var loginPageResponse = UrlFetchApp.fetch(loginUrl, loginPageOptions);
//     var loginPageHtml = loginPageResponse.getContentText("UTF-8");
//     var cookies = loginPageResponse.getAllHeaders()["Set-Cookie"];

//     // CSRFトークンを抽出（フォーム内の_tokenまたはmeta tagから）
//     var csrfToken = "";
//     var csrfMatch = loginPageHtml.match(
//       /<input[^>]+name="_token"[^>]+value="([^"]+)"/
//     );
//     if (!csrfMatch) {
//       csrfMatch = loginPageHtml.match(
//         /<meta name="csrf-token" content="([^"]+)"/
//       );
//     }
//     if (csrfMatch) {
//       csrfToken = csrfMatch[1];
//     }

//     console.log("CSRFトークン取得:", csrfToken ? "成功" : "失敗");

//     // ログインフォームのアクションURLを確認
//     var formActionMatch = loginPageHtml.match(
//       /<form[^>]+action="([^"]+)"[^>]*>/
//     );
//     var actualLoginUrl = loginUrl;
//     if (formActionMatch) {
//       var action = formActionMatch[1];
//       if (action.startsWith("/")) {
//         actualLoginUrl = "https://www.starbuyers-global-auction.com" + action;
//       } else if (action.startsWith("http")) {
//         actualLoginUrl = action;
//       }
//     }
//     console.log("実際のログインURL:", actualLoginUrl);

//     // ログインフォームのフィールドを確認
//     var emailFieldMatch = loginPageHtml.match(
//       /<input[^>]+type="email"[^>]+name="([^"]+)"/
//     );
//     var emailFieldName = emailFieldMatch ? emailFieldMatch[1] : "email";
//     console.log("Emailフィールド名:", emailFieldName);

//     // セッションクッキーを収集
//     var cookieMap = {};
//     if (cookies) {
//       if (Array.isArray(cookies)) {
//         cookies.forEach(function (cookie) {
//           var parts = cookie.split(";")[0].split("=");
//           if (parts.length >= 2) {
//             cookieMap[parts[0]] = parts.slice(1).join("=");
//           }
//         });
//       } else {
//         var parts = cookies.split(";")[0].split("=");
//         if (parts.length >= 2) {
//           cookieMap[parts[0]] = parts.slice(1).join("=");
//         }
//       }
//     }

//     console.log("取得したクッキー:", Object.keys(cookieMap).join(", "));

//     // ランダムな待機時間（1-2秒）
//     Utilities.sleep(1000 + Math.floor(Math.random() * 1000));

//     // ステップ2: ログイン実行
//     var loginPayload = {};
//     loginPayload[emailFieldName] = loginEmail;
//     loginPayload["password"] = loginPassword;
//     loginPayload["_token"] = csrfToken;

//     // Remember meチェックボックスがある場合
//     if (loginPageHtml.indexOf('name="remember"') > -1) {
//       loginPayload["remember"] = "1";
//     }

//     // クッキー文字列を構築
//     var cookieString = Object.keys(cookieMap)
//       .map(function (key) {
//         return key + "=" + cookieMap[key];
//       })
//       .join("; ");

//     // XSRF-TOKENがある場合、X-XSRF-TOKENヘッダーとして追加
//     var xsrfToken = cookieMap["XSRF-TOKEN"] || "";
//     if (xsrfToken) {
//       // URLデコード（LaravelのXSRF-TOKENは通常URLエンコードされている）
//       try {
//         xsrfToken = decodeURIComponent(xsrfToken);
//       } catch (e) {}
//     }

//     var loginHeaders = {
//       "User-Agent":
//         "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//       Accept:
//         "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//       "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//       "Content-Type": "application/x-www-form-urlencoded",
//       Cookie: cookieString,
//       Referer: loginUrl,
//       Origin: "https://www.starbuyers-global-auction.com",
//       "sec-ch-ua":
//         '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
//       "sec-ch-ua-mobile": "?0",
//       "sec-ch-ua-platform": '"macOS"',
//       "sec-fetch-dest": "document",
//       "sec-fetch-mode": "navigate",
//       "sec-fetch-site": "same-origin",
//       "sec-fetch-user": "?1",
//       "upgrade-insecure-requests": "1",
//     };

//     if (xsrfToken) {
//       loginHeaders["X-XSRF-TOKEN"] = xsrfToken;
//     }

//     var payloadString = Object.keys(loginPayload)
//       .map(function (key) {
//         return (
//           encodeURIComponent(key) + "=" + encodeURIComponent(loginPayload[key])
//         );
//       })
//       .join("&");

//     console.log(
//       "ログインペイロード:",
//       payloadString.replace(loginPassword, "***")
//     );

//     var loginOptions = {
//       method: "post",
//       headers: loginHeaders,
//       payload: payloadString,
//       muteHttpExceptions: true,
//       followRedirects: false,
//     };

//     var loginResponse = UrlFetchApp.fetch(actualLoginUrl, loginOptions);
//     console.log(
//       "ログインレスポンスステータス:",
//       loginResponse.getResponseCode()
//     );

//     // ログイン後のクッキーを更新
//     var loginCookies = loginResponse.getAllHeaders()["Set-Cookie"];
//     if (loginCookies) {
//       if (Array.isArray(loginCookies)) {
//         loginCookies.forEach(function (cookie) {
//           var parts = cookie.split(";")[0].split("=");
//           if (parts.length >= 2) {
//             cookieMap[parts[0]] = parts.slice(1).join("=");
//           }
//         });
//       } else {
//         var parts = loginCookies.split(";")[0].split("=");
//         if (parts.length >= 2) {
//           cookieMap[parts[0]] = parts.slice(1).join("=");
//         }
//       }
//     }

//     // リダイレクト先を確認
//     var locationHeader =
//       loginResponse.getAllHeaders()["Location"] ||
//       loginResponse.getAllHeaders()["location"];
//     console.log("リダイレクト先:", locationHeader || "なし");
//     if (locationHeader) {
//       // リダイレクト先にアクセス
//       cookieString = Object.keys(cookieMap)
//         .map(function (key) {
//           return key + "=" + cookieMap[key];
//         })
//         .join("; ");

//       var redirectOptions = {
//         method: "get",
//         headers: {
//           "User-Agent":
//             "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//           Accept:
//             "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//           Cookie: cookieString,
//         },
//         muteHttpExceptions: true,
//       };

//       var redirectResponse = UrlFetchApp.fetch(locationHeader, redirectOptions);
//       var redirectCookies = redirectResponse.getAllHeaders()["Set-Cookie"];
//       if (redirectCookies) {
//         if (Array.isArray(redirectCookies)) {
//           redirectCookies.forEach(function (cookie) {
//             var parts = cookie.split(";")[0].split("=");
//             if (parts.length >= 2) {
//               cookieMap[parts[0]] = parts.slice(1).join("=");
//             }
//           });
//         }
//       }
//     }

//     // ランダムな待機時間（2-4秒）
//     Utilities.sleep(2000 + Math.floor(Math.random() * 2000));

//     // ステップ3: 目的のURLにアクセス
//     cookieString = Object.keys(cookieMap)
//       .map(function (key) {
//         return key + "=" + cookieMap[key];
//       })
//       .join("; ");

//     var fetchOptions = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//         Cookie: cookieString,
//         Referer: "https://www.starbuyers-global-auction.com/",
//         "sec-ch-ua":
//           '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
//         "sec-ch-ua-mobile": "?0",
//         "sec-ch-ua-platform": '"macOS"',
//         "sec-fetch-dest": "document",
//         "sec-fetch-mode": "navigate",
//         "sec-fetch-site": "none",
//         "sec-fetch-user": "?1",
//         "upgrade-insecure-requests": "1",
//       },
//       muteHttpExceptions: true,
//     };

//     var response = UrlFetchApp.fetch(url, fetchOptions);
//     var html = response.getContentText("UTF-8");

//     if (response.getResponseCode() !== 200) {
//       throw new Error(
//         "Star Buyersからのデータ取得に失敗しました。ステータスコード: " +
//           response.getResponseCode()
//       );
//     }

//     // ログインが必要なページにリダイレクトされていないかチェック
//     if (
//       html.indexOf("Login") > -1 &&
//       html.indexOf("E-Mail Address") > -1 &&
//       html.indexOf("p-item-list__body") === -1
//     ) {
//       console.log("HTMLレスポンスの一部:", html.substring(0, 500));
//       throw new Error("ログインに失敗しました。認証情報を確認してください。");
//     }

//     // 商品データが含まれているかチェック
//     if (html.indexOf("p-item-list__body") === -1) {
//       console.log("取得したHTMLに商品データが含まれていない可能性があります。");
//     }

//     return html;
//   } catch (e) {
//     throw new Error("Star Buyers URLからのHTMLフェッチエラー: " + e.message);
//   }
// }

// // オークファンURLからHTMLをフェッチ（未ログイン前提の一覧/詳細を対象）
// function fetchAucfanHtml_(url) {
//   try {
//     var options = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//       },
//       muteHttpExceptions: true,
//       followRedirects: true,
//     };

//     var response = UrlFetchApp.fetch(url, options);
//     var html = response.getContentText("UTF-8");
//     if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
//       throw new Error(
//         "オークファンからのHTMLの取得に失敗しました。ステータスコード: " +
//           response.getResponseCode()
//       );
//     }
//     return html;
//   } catch (e) {
//     throw new Error("オークファンURLからのHTMLフェッチエラー: " + e.message);
//   }
// }

// // エコリングURLからHTMLをフェッチ（ログイン必要）
// function fetchEcoringHtml_(url) {
//   try {
//     logInfo_("fetchEcoringHtml_", "エコリングHTML取得開始", "URL: " + url);
//     // ランダムな待機時間（1-3秒）
//     Utilities.sleep(1000 + Math.floor(Math.random() * 2000));

//     // ログイン情報（ベタ打ち）
//     var loginEmail = "info@genkaya.jp";
//     var loginPassword = "ecoringenkaya";

//     // ログインURL
//     var loginUrl = "https://www.ecoauc.com/client/users/sign-in";

//     // ステップ1: ログインページにアクセスしてクッキーを取得
//     var loginPageOptions = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//       },
//       muteHttpExceptions: true,
//       followRedirects: false,
//     };

//     var loginPageResponse = UrlFetchApp.fetch(loginUrl, loginPageOptions);
//     var loginPageHtml = loginPageResponse.getContentText("UTF-8");
//     var cookies = loginPageResponse.getAllHeaders()["Set-Cookie"];

//     // CSRFトークンを抽出（エコリングは_csrfTokenを使用）
//     var csrfToken = "";
//     var csrfMatch = loginPageHtml.match(
//       /<input[^>]+name="_csrfToken"[^>]+value="([^"]+)"/
//     );
//     if (!csrfMatch) {
//       // 属性の順序が異なる場合
//       csrfMatch = loginPageHtml.match(
//         /<input[^>]+value="([^"]+)"[^>]+name="_csrfToken"/
//       );
//     }
//     if (csrfMatch) {
//       csrfToken = csrfMatch[1];
//     }

//     console.log("エコリング CSRFトークン取得:", csrfToken ? "成功" : "失敗");
//     if (!csrfToken) {
//       logWarning_("fetchEcoringHtml_", "CSRFトークンが取得できませんでした", "");
//     }

//     // セッションクッキーを収集
//     var cookieMap = {};
//     if (cookies) {
//       if (Array.isArray(cookies)) {
//         cookies.forEach(function (cookie) {
//           var parts = cookie.split(";")[0].split("=");
//           if (parts.length >= 2) {
//             cookieMap[parts[0]] = parts.slice(1).join("=");
//           }
//         });
//       } else {
//         var parts = cookies.split(";")[0].split("=");
//         if (parts.length >= 2) {
//           cookieMap[parts[0]] = parts.slice(1).join("=");
//         }
//       }
//     }

//     console.log("エコリング 取得したクッキー:", Object.keys(cookieMap).join(", "));

//     // ランダムな待機時間（1-2秒）
//     Utilities.sleep(1000 + Math.floor(Math.random() * 1000));

//     // ステップ2: ログイン実行（エコリングのフォームに合わせる）
//     var loginPayload = {
//       "_method": "POST",
//       "_csrfToken": csrfToken,
//       "email_address": loginEmail,
//       "password": loginPassword
//     };

//     // クッキー文字列を構築
//     var cookieString = Object.keys(cookieMap)
//       .map(function (key) {
//         return key + "=" + cookieMap[key];
//       })
//       .join("; ");

//     var loginHeaders = {
//       "User-Agent":
//         "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//       Accept:
//         "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//       "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//       "Content-Type": "application/x-www-form-urlencoded",
//       Cookie: cookieString,
//       Referer: loginUrl,
//       Origin: "https://www.ecoauc.com",
//     };

//     var payloadString = Object.keys(loginPayload)
//       .map(function (key) {
//         return (
//           encodeURIComponent(key) + "=" + encodeURIComponent(loginPayload[key])
//         );
//       })
//       .join("&");

//     console.log(
//       "エコリング ログインペイロード:",
//       payloadString.replace(loginPassword, "***")
//     );

//     var loginOptions = {
//       method: "post",
//       headers: loginHeaders,
//       payload: payloadString,
//       muteHttpExceptions: true,
//       followRedirects: false,
//     };

//     // ログインURLを正しいエンドポイントに変更
//     var loginPostUrl = "https://www.ecoauc.com/client/users/post-sign-in";
//     var loginResponse = UrlFetchApp.fetch(loginPostUrl, loginOptions);
//     console.log(
//       "エコリング ログインレスポンスステータス:",
//       loginResponse.getResponseCode()
//     );
    
//     // ログインエラーの場合、レスポンス内容を確認
//     if (loginResponse.getResponseCode() !== 302 && loginResponse.getResponseCode() !== 200) {
//       var loginErrorHtml = loginResponse.getContentText("UTF-8");
//       console.log("エコリング ログインエラーレスポンス（最初の1000文字）:", loginErrorHtml.substring(0, 1000));
//       logError_("fetchEcoringHtml_", "ERROR", "ログインエラー", "ステータス: " + loginResponse.getResponseCode() + ", 内容: " + loginErrorHtml.substring(0, 500));
//     }

//     // ログイン後のクッキーを更新
//     var loginCookies = loginResponse.getAllHeaders()["Set-Cookie"];
//     if (loginCookies) {
//       if (Array.isArray(loginCookies)) {
//         loginCookies.forEach(function (cookie) {
//           var parts = cookie.split(";")[0].split("=");
//           if (parts.length >= 2) {
//             cookieMap[parts[0]] = parts.slice(1).join("=");
//           }
//         });
//       } else {
//         var parts = loginCookies.split(";")[0].split("=");
//         if (parts.length >= 2) {
//           cookieMap[parts[0]] = parts.slice(1).join("=");
//         }
//       }
//     }

//     // リダイレクト先を確認
//     var locationHeader =
//       loginResponse.getAllHeaders()["Location"] ||
//       loginResponse.getAllHeaders()["location"];
//     console.log("エコリング リダイレクト先:", locationHeader || "なし");
    
//     if (locationHeader) {
//       // リダイレクト先にアクセス
//       cookieString = Object.keys(cookieMap)
//         .map(function (key) {
//           return key + "=" + cookieMap[key];
//         })
//         .join("; ");

//       var redirectOptions = {
//         method: "get",
//         headers: {
//           "User-Agent":
//             "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//           Accept:
//             "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//           Cookie: cookieString,
//         },
//         muteHttpExceptions: true,
//       };

//       var redirectResponse = UrlFetchApp.fetch(locationHeader, redirectOptions);
//       var redirectCookies = redirectResponse.getAllHeaders()["Set-Cookie"];
//       if (redirectCookies) {
//         if (Array.isArray(redirectCookies)) {
//           redirectCookies.forEach(function (cookie) {
//             var parts = cookie.split(";")[0].split("=");
//             if (parts.length >= 2) {
//               cookieMap[parts[0]] = parts.slice(1).join("=");
//             }
//           });
//         }
//       }
//     }

//     // ランダムな待機時間（2-4秒）
//     Utilities.sleep(2000 + Math.floor(Math.random() * 2000));

//     // ステップ3: 目的のURLにアクセス
//     cookieString = Object.keys(cookieMap)
//       .map(function (key) {
//         return key + "=" + cookieMap[key];
//       })
//       .join("; ");

//     var fetchOptions = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//         Cookie: cookieString,
//         Referer: "https://www.ecoauc.com/",
//       },
//       muteHttpExceptions: true,
//     };

//     var response = UrlFetchApp.fetch(url, fetchOptions);
//     var html = response.getContentText("UTF-8");

//     if (response.getResponseCode() !== 200) {
//       logError_("fetchEcoringHtml_", "ERROR", "HTTPステータスエラー", "ステータスコード: " + response.getResponseCode());
//       throw new Error(
//         "エコリングからのデータ取得に失敗しました。ステータスコード: " +
//           response.getResponseCode()
//       );
//     }

//     // ログインが必要なページにリダイレクトされていないかチェック
//     if (
//       html.indexOf("sign-in") > -1 &&
//       html.indexOf("ログイン") > -1 &&
//       html.indexOf("market-prices") === -1
//     ) {
//       console.log("エコリング HTMLレスポンスの一部:", html.substring(0, 500));
//       logError_("fetchEcoringHtml_", "ERROR", "ログイン失敗", "ログインページにリダイレクトされました。認証情報を確認してください。");
//       throw new Error("エコリング ログインに失敗しました。認証情報を確認してください。");
//     }
    
//     // 商品データが含まれているかチェック
//     if (html.indexOf("show-case-title-block") === -1) {
//       console.log("エコリング 取得したHTMLに商品データが含まれていない可能性があります。");
//       console.log("エコリング HTMLの最初の1500文字:", html.substring(0, 1500));
//       logWarning_("fetchEcoringHtml_", "商品データなし", "HTMLに商品データが見つかりません。HTMLサイズ: " + html.length);
//     }

//     logInfo_("fetchEcoringHtml_", "HTML取得成功", "HTMLサイズ: " + html.length);
//     return html;
//   } catch (e) {
//     logError_("fetchEcoringHtml_", "ERROR", "HTMLフェッチエラー", e.message);
//     throw new Error("エコリング URLからのHTMLフェッチエラー: " + e.message);
//   }
// }

// // HTMLエンティティ最小限デコード
// function htmlDecode_(s) {
//   if (!s) return s;
//   return s
//     .replace(/&amp;/g, "&")
//     .replace(/&lt;/g, "<")
//     .replace(/&gt;/g, ">")
//     .replace(/&quot;/g, '"')
//     .replace(/&#039;/g, "'")
//     .replace(/&nbsp;/g, " ");
// }

// // レスポンスの文字コードを検出して適切にテキスト化
// function detectCharsetFromResponse_(response) {
//   try {
//     var headers = response.getAllHeaders() || {};
//     var ct = headers["Content-Type"] || headers["content-type"] || "";
//     var m = String(ct).match(/charset=([^;]+)/i);
//     var headerCs = m ? (m[1] || "").trim() : "";

//     // 本文から meta charset を検出（ASCII として走査）
//     var bytes = response.getContent();
//     var ascii = Utilities.newBlob(bytes).getDataAsString("ISO-8859-1");
//     var metaCs = "";
//     var m1 = ascii.match(/<meta[^>]+charset\s*=\s*"?([A-Za-z0-9_\-]+)"?/i);
//     if (m1) metaCs = m1[1];
//     if (!metaCs) {
//       var m2 = ascii.match(
//         /<meta[^>]+http-equiv=["']?Content-Type["']?[^>]+content=["'][^"']*charset=([A-Za-z0-9_\-]+)["']/i
//       );
//       if (m2) metaCs = m2[1];
//     }

//     var cs = headerCs || metaCs || "";
//     cs = cs.toUpperCase();
//     if (
//       cs.indexOf("SHIFT") !== -1 ||
//       cs.indexOf("SJIS") !== -1 ||
//       cs.indexOf("MS932") !== -1
//     )
//       return "Shift_JIS";
//     if (cs.indexOf("EUC") !== -1) return "EUC-JP";
//     if (cs.indexOf("UTF-8") !== -1 || cs.indexOf("UTF8") !== -1) return "UTF-8";
//     return cs || "UTF-8";
//   } catch (e) {
//     return "UTF-8";
//   }
// }

// function getResponseTextWithBestCharset_(response) {
//   var charset = detectCharsetFromResponse_(response) || "UTF-8";
//   try {
//     return response.getContentText(charset);
//   } catch (e) {
//     return response.getContentText(); // 既定 (UTF-8)
//   }
// }

// // タグ除去（簡易）
// function stripTags_(s) {
//   if (!s) return "";
//   return String(s)
//     .replace(/<[^>]*>/g, "")
//     .trim();
// }

// // C列以降を対象行だけクリア
// function clearRowFromColumnC_(sheet, row) {
//   var lastCol = sheet.getLastColumn();
//   if (lastCol >= COL_C) {
//     sheet.getRange(row, COL_C, 1, lastCol - (COL_C - 1)).clearContent();
//   }
// }

// // B列の全見出し（1〜lastRow）を配列で取得（trim）
// function readColumnBHeadings_(sheet) {
//   var last = Math.max(sheet.getLastRow(), 1000);
//   var bVals = sheet.getRange(1, COL_B, last, 1).getValues();
//   return bVals.map(function (r) {
//     return String(r[0] || "").trim();
//   });
// }

// // ラベル→横並びデータを書き込み（C列起点）
// function writeRowByLabel_(sheet, bHeadings, label, rowValuesArray) {
//   var idx = bHeadings.findIndex(function (v) {
//     return v === label;
//   });
//   if (idx === -1) return; // 見出し未設置ならスキップ
//   var row = idx + 1;
//   clearRowFromColumnC_(sheet, row);
//   if (!rowValuesArray || !rowValuesArray.length) return;

//   // 件数制限（C〜Vの20件以内）
//   var out = rowValuesArray.slice(0, MAX_OUTPUT_ITEMS);
//   var values2d = [out];
//   sheet.getRange(row, COL_C, 1, out.length).setValues(values2d);
// }

// /** ============== ソース判定 ============== **/
// function detectSource_(html) {
//   var h = html.toLowerCase();
//   if (
//     h.indexOf("starbuyers-global-auction.com") !== -1 ||
//     h.indexOf("star buyers auction") !== -1 ||
//     h.indexOf("p-item-list__body__cell") !== -1
//   ) {
//     return "starbuyers";
//   }
//   if (
//     h.indexOf("aucfan") !== -1 ||
//     h.indexOf("オークファン") !== -1 ||
//     h.indexOf("落札相場") !== -1
//   ) {
//     return "aucfan";
//   }
//   if (
//     h.indexOf("rakuten.co.jp") !== -1 ||
//     h.indexOf("楽天市場") !== -1 ||
//     h.indexOf("送料無料") !== -1 ||
//     h.indexOf("ポイント(1倍)") !== -1
//   ) {
//     return "rakuten";
//   }
//   if (
//     h.indexOf("auctions.yahoo.co.jp") !== -1 ||
//     h.indexOf("ヤフオク") !== -1 ||
//     h.indexOf("落札価格") !== -1 ||
//     h.indexOf("入札件数") !== -1
//   ) {
//     return "yahooauction";
//   }
//   return "starbuyers";
// }
// // --- 置き換え版：STAR BUYERS パーサ ---
// // 返却：[{detailUrl, imageUrl, date, rank, price}]
// function parseStarBuyersFromHtml_(html) {
//   var items = [];
//   var H = String(html);

//   console.log("SBAパーサー - HTMLの長さ:", H.length);
//   console.log(
//     "SBAパーサー - HTMLに'p-item-list__body'が含まれているか:",
//     H.indexOf("p-item-list__body") > -1
//   );
//   console.log(
//     "SBAパーサー - HTMLに'market_price'が含まれているか:",
//     H.indexOf("market_price") > -1
//   );

//   // 1) 各アイテムの開始インデックスを列挙
//   var starts = [];
//   var re = /<div\s+class="p-item-list__body">/g;
//   var m;
//   while ((m = re.exec(H)) !== null) {
//     starts.push(m.index);
//   }
//   console.log("SBAパーサー - p-item-list__bodyが見つかった数:", starts.length);
//   if (starts.length === 0) return items;

//   // 2) 開始インデックスごとに「次の開始 or 文末」までを1件ブロックとする
//   for (var i = 0; i < starts.length; i++) {
//     var s = starts[i];
//     var e = i + 1 < starts.length ? starts[i + 1] : H.length;
//     var block = H.slice(s, e);

//     // 詳細URL
//     var detailUrl = "";
//     var mUrl = block.match(
//       /<a[^>]+href="(https:\/\/www\.starbuyers-global-auction\.com\/market_price\/[^"]+)"/i
//     );
//     if (mUrl) detailUrl = htmlDecode_(mUrl[1]);

//     // 画像URL
//     var imageUrl = "";
//     var mImg = block.match(/<img[^>]+src="([^"]+)"[^>]*>/i);
//     if (mImg) imageUrl = htmlDecode_(mImg[1]);

//     // 商品名（名前セル内の p-text-link を優先して抽出。画像側のアンカーは除外）
//     var title = "";
//     var nameCell = firstMatch_(
//       block,
//       /<p[^>]+class="[^"]*p-item-list__body__cell[^"]* -name[^"]*"[^>]*>([\s\S]*?)<\/p>/i
//     );
//     if (nameCell) {
//       var mNameLink = nameCell.match(
//         /<a[^>]+class="[^"]*p-text-link[^"]*"[^>]*>([\s\S]*?)<\/a>/i
//       );
//       if (mNameLink) {
//         title = stripTags_(htmlDecode_(mNameLink[1]))
//           .replace(/\s+/g, " ")
//           .trim();
//       }
//       if (!title) {
//         var mAnyNameLink = nameCell.match(
//           /<a[^>]+href="https:\/\/www\.starbuyers-global-auction\.com\/market_price\/[^"]+"[^>]*>([\s\S]*?)<\/a>/i
//         );
//         if (mAnyNameLink) {
//           title = stripTags_(htmlDecode_(mAnyNameLink[1]))
//             .replace(/\s+/g, " ")
//             .trim();
//         }
//       }
//     }
//     if (!title) {
//       var mTitleHead = block.match(
//         /data-head="商品名"[\s\S]*?<strong>([^<]+)<\/strong>/i
//       );
//       if (mTitleHead) title = mTitleHead[1].trim();
//     }

//     // 開催日
//     var date = "";
//     var mDate = block.match(
//       /data-head="開催日"[\s\S]*?<strong>([^<]+)<\/strong>/i
//     );
//     if (mDate) date = mDate[1].trim();

//     // ランク（data-rank="ＡＢ" など）→ 許可値に正規化
//     var rank = "";
//     var mRank = block.match(/data-rank="([^"]+)"/i);
//     if (mRank) rank = normalizeRank_(mRank[1].trim());

//     // 価格（"96,000yen" 等）
//     var price = "";
//     var mPrice = block.match(
//       /data-head="落札額"[\s\S]*?<strong>([^<]+)<\/strong>/i
//     );
//     if (mPrice) {
//       var raw = mPrice[1].trim();
//       var num = normalizeNumber_(raw.replace(/yen/i, ""));
//       // 数値化できた場合は数値、できなかった場合は元のテキストを保持
//       price = num || raw;
//       console.log("SBA価格抽出:", raw, "→", price);
//     }

//     if (detailUrl || imageUrl || date || rank || price || title) {
//       items.push({
//         title: title,
//         detailUrl: detailUrl,
//         imageUrl: imageUrl,
//         date: date,
//         rank: rank,
//         price: price,
//         shop: "",
//         source: "SBA",
//       });
//     }
//   }

//   return items;
// }

// /** オークファン用パーサ（元の extractAucfanItems を簡略化して利用） */
// function parseAucfanFromHtml_(html) {
//   const items = [];
//   const base = "https://pro.aucfan.com";
//   const re = /<li class="item">\s*<ul>([\s\S]*?)<\/ul>\s*<\/li>/gi;
//   let m;
//   while ((m = re.exec(html)) !== null) {
//     const block = m[1];
//     const detailPath = firstMatch_(
//       block,
//       /<li class="title">[\s\S]*?<a\s+href="([^"]+)"/i
//     );
//     const detailUrl = detailPath
//       ? detailPath.startsWith("http")
//         ? detailPath
//         : base + detailPath
//       : "";
//     const imageUrl = firstMatch_(
//       block,
//       /<img[^>]+class="thumbnail"[^>]+src="([^"]+)"/i
//     );
//     // タイトル抽出（itemName のみ使用）
//     let title = "";
//     const titleHtmlItemName = firstMatch_(
//       block,
//       /<li class="itemName">[\s\S]*?<a[^>]*>([\s\S]*?)<\/a>/i
//     );
//     if (titleHtmlItemName) {
//       title = stripTags_(htmlDecode_(titleHtmlItemName))
//         .replace(/\s+/g, " ")
//         .trim();
//     }
//     if (!title) {
//       const liItemNameAll = firstMatch_(
//         block,
//         /<li class="itemName">([\s\S]*?)<\/li>/i
//       );
//       if (liItemNameAll) {
//         title = stripTags_(htmlDecode_(liItemNameAll))
//           .replace(/\s+/g, " ")
//           .trim();
//       }
//     }
//     let price = "";
    
//     // まず落札価格を探す（落札価格を優先）
//     // パターン1: <span>落札</span>の後の価格
//     let endPrice = firstMatch_(
//       block,
//       /<li class="price"[^>]*>\s*<span[^>]*>落札<\/span>\s*([^<]+)/i
//     );
    
//     // パターン2: 落札価格として明記されている場合
//     if (!endPrice) {
//       endPrice = firstMatch_(
//         block,
//         /落札価格[:：]?\s*([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i
//       );
//     }
    
//     // パターン3: 終了価格として表示されている場合
//     if (!endPrice) {
//       endPrice = firstMatch_(
//         block,
//         /終了価格[:：]?\s*([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i
//       );
//     }
    
//     if (endPrice) {
//       // 落札価格が見つかった場合はそれを使用
//       const normalizedEndPrice = normalizeNumber_(endPrice);
//       if (normalizedEndPrice && normalizedEndPrice !== "0") {
//         price = normalizedEndPrice;
//         console.log("オークファン - 落札価格を使用:", price);
//       }
//     }
    
//     // 落札価格が見つからない場合は、既存の価格抽出ロジックを使用
//     if (!price) {
//       const priceCandidates = [];
//       const pricePrimary = firstMatch_(
//         block,
//         /<li class="price"[^>]*>\s*([^<]+)/i
//       );
//       if (pricePrimary) priceCandidates.push(pricePrimary);
//       const dataPrice = firstMatch_(block, /data-price="([0-9,]+)"/i);
//       if (dataPrice) priceCandidates.push(dataPrice);
//       const priceValue = firstMatch_(
//         block,
//         /class="[^"]*price__value[^"]*"[^>]*>\s*([^<]+)/i
//       );
//       if (priceValue) priceCandidates.push(priceValue);
//       const yenInline = firstMatch_(block, /([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i);
//       if (yenInline) priceCandidates.push(yenInline);
//       const yenMark = firstMatch_(block, /¥\s*([0-9]{1,3}(?:,[0-9]{3})*)/i);
//       if (yenMark) priceCandidates.push(yenMark);

//       const nums = priceCandidates
//         .map(function (p) {
//           return normalizeNumber_(p);
//         })
//         .filter(function (n) {
//           return n && n !== "0";
//         })
//         .map(function (n) {
//           return parseInt(n, 10);
//         })
//         .filter(function (n) {
//           return !isNaN(n) && n >= 100;
//         });
//       if (nums.length > 0) {
//         var filtered = nums.filter(function (n) {
//           return (
//             [
//               3980, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000,
//             ].indexOf(n) === -1
//           );
//         });
//         var candidates = filtered.length > 0 ? filtered : nums;
//         candidates.sort(function (a, b) {
//           return a - b;
//         });
//         price = String(candidates[0]);
//       }
//     }
//     const endTxt = firstMatch_(block, /<li class="end">\s*([^<]+)/i);
//     // endTxtそのまま日付に入れる（必要なら deriveEndDateValue_ を再利用可）

//     if (detailUrl || imageUrl || price || endTxt || title) {
//       items.push({
//         title: title,
//         detailUrl: detailUrl,
//         imageUrl: imageUrl,
//         date: endTxt || "",
//         rank: "", // aucfanにはランクが基本出ないので空
//         price: price || "",
//         shop: "",
//         source: "オークファン",
//       });
//     }
//   }
//   return items;
// }

// /** 楽天用パーサ */
// function parseRakutenFromHtml_(html) {
//   const items = [];
//   const H = String(html);

//   // 方法1: 商品ブロックを特定のパターンで抽出を試みる
//   // 楽天の商品は通常、特定のクラスを持つdivで囲まれている
//   const blockPatterns = [
//     // 楽天の新レイアウト: dui-cardクラス
//     /<div[^>]+class="[^"]*dui-card[^"]*"[^>]*>[\s\S]*?(?=<div[^>]+class="[^"]*dui-card|$)/gi,
//     // 楽天の検索結果の商品全体を含むブロック（より広範囲に取得）
//     /<div[^>]+class="[^"]*searchresultitem[^"]*"[^>]*>[\s\S]*?(?=<div[^>]+class="[^"]*searchresultitem|$)/gi,
//     // 商品リンクから次の商品リンクまでのブロック（範囲を広げる）
//     /<a[^>]+href="https?:\/\/item\.rakuten\.co\.jp\/[^"]+[^>]*>[\s\S]{0,5000}?(?=<a[^>]+href="https?:\/\/item\.rakuten\.co\.jp\/|$)/gi,
//     // content-bottomを含むブロック
//     /<div[^>]+class="[^"]*content[^"]*"[^>]*>[\s\S]*?<div[^>]+class="[^"]*content-bottom[^"]*"[^>]*>[\s\S]*?<\/div>/gi,
//     // importantとpriceを含むブロック
//     /<div[^>]+class="[^"]*important[^"]*"[^>]*>[\s\S]*?<\/div>[\s\S]*?<div[^>]+class="[^"]*price[^"]*"[^>]*>[\s\S]*?<\/div>/gi,
//     // 商品を含むliタグ
//     /<li[^>]*>[\s\S]*?<a[^>]+href="https?:\/\/item\.rakuten\.co\.jp[^>]*>[\s\S]*?<\/li>/gi,
//   ];

//   let productBlocks = [];

//   // 各パターンを試して商品ブロックを抽出
//   for (let pattern of blockPatterns) {
//     const matches = H.match(pattern);
//     if (matches && matches.length > 0) {
//       productBlocks = matches;
//       console.log(
//         "楽天商品ブロック検出:",
//         matches.length + "件（パターン:",
//         blockPatterns.indexOf(pattern) + "）"
//       );
//       break;
//     }
//   }

//   // 商品ブロックが見つかった場合
//   if (productBlocks.length > 0) {
//     let useBlockMethod = true;
//     const testItems = [];

//     // 最初の3つの商品ブロックをテストして、価格が抽出できるか確認
//     for (let i = 0; i < Math.min(3, productBlocks.length); i++) {
//       const block = productBlocks[i];
//       const detailUrlMatch = block.match(
//         /<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/i
//       );
//       if (!detailUrlMatch) continue;

//       const detailUrl = htmlDecode_(detailUrlMatch[1]);
//       const testItem = extractRakutenItemFromBlock_(block, detailUrl, i);
//       testItems.push(testItem);

//       // 価格が見つからない場合
//       if (!testItem.price) {
//         console.log(
//           "ブロック方式で価格が見つかりません。リンクベース方式に切り替えます。"
//         );
//         useBlockMethod = false;
//         break;
//       }
//     }

//     if (useBlockMethod) {
//       // ブロック方式で全商品を処理
//       productBlocks.forEach(function (block, index) {
//         const detailUrlMatch = block.match(
//           /<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/i
//         );
//         if (!detailUrlMatch) return;

//         const detailUrl = htmlDecode_(detailUrlMatch[1]);
//         items.push(extractRakutenItemFromBlock_(block, detailUrl, index));
//       });
//     } else {
//       // リンクベース方式にフォールバック
//       productBlocks = [];
//     }
//   }

//   if (productBlocks.length === 0) {
//     // ブロックが見つからない場合は、従来の方法（リンクベース）を使用
//     console.log(
//       "商品ブロックが見つかりません。リンクベースの抽出に切り替えます。"
//     );
//     const linkRe =
//       /<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"[^>]*>/gi;
//     let links = [];
//     let m;
//     while ((m = linkRe.exec(H)) !== null) {
//       links.push({
//         url: m[1],
//         startIndex: m.index,
//         endIndex: m.index + m[0].length,
//       });
//     }

//     // 各リンクに対して、周辺の情報を抽出
//     links.forEach(function (link, index) {
//       // リンクの前後2500文字を検索範囲とする（さらに範囲を広げる）
//       const searchStart = Math.max(0, link.startIndex - 2500);
//       const searchEnd = Math.min(H.length, link.endIndex + 2500);
//       const block = H.substring(searchStart, searchEnd);

//       const detailUrl = htmlDecode_(link.url);
//       items.push(extractRakutenItemFromBlock_(block, detailUrl, index));
//     });
//   }

//   return items;
// }

// // 楽天の商品ブロックから情報を抽出するヘルパー関数
// function extractRakutenItemFromBlock_(block, detailUrl, index) {
//   // デバッグ：最初の3商品のブロックをログ出力
//   if (index < 3) {
//     console.log("=== 楽天商品ブロック " + (index + 1) + " ===");
//     console.log("URL:", detailUrl);
//     console.log("ブロック長:", block.length);

//     // 実際の商品価格を探すためのデバッグ
//     const allPriceMatches = block.match(/[0-9,]+\s*円/g);
//     if (allPriceMatches) {
//       console.log(
//         "ブロック内の全価格候補:",
//         allPriceMatches.filter((p) => !p.includes("3,980")).slice(0, 5)
//       );
//     }

//     // 価格タグを探す
//     const priceTagMatches = block.match(
//       /<(?:span|div|p)[^>]*(?:class|id)="[^"]*price[^"]*"[^>]*>[^<]+<\/(?:span|div|p)>/gi
//     );
//     if (priceTagMatches) {
//       console.log("価格タグ数:", priceTagMatches.length);
//       if (priceTagMatches[0]) {
//         console.log("最初の価格タグ:", priceTagMatches[0].replace(/\s+/g, " "));
//       }
//     }

//     // 商品リンクの後の価格を探す
//     const afterLinkMatch = block.match(
//       /item\.rakuten\.co\.jp[^>]+>[\s\S]{0,300}?([0-9,]+)\s*円/
//     );
//     if (afterLinkMatch) {
//       console.log("商品リンク後の価格:", afterLinkMatch[1] + "円");
//     }
//   }

//   // 商品名の抽出
//   let title = "";
//   // 1) 対象リンクのアンカーテキスト
//   const anchorTitle = firstMatch_(
//     block,
//     /<a[^>]+href="https?:\/\/item\.rakuten\.co\.jp\/[^"]+"[^>]*>([\s\S]*?)<\/a>/i
//   );
//   if (anchorTitle) {
//     title = stripTags_(htmlDecode_(anchorTitle)).replace(/\s+/g, " ").trim();
//   }
//   // 2) img の alt 属性
//   if (!title) {
//     const altTitle = firstMatch_(block, /<img[^>]+alt="([^"]+)"[^>]*>/i);
//     if (altTitle) {
//       title = htmlDecode_(altTitle).trim();
//     }
//   }
//   // 3) タイトル系クラス
//   if (!title) {
//     const classTitle = firstMatch_(
//       block,
//       /<(?:div|span|p)[^>]*class="[^"]*(?:title|productName|content-title)[^"]*"[^>]*>([\s\S]*?)<\/(?:div|span|p)>/i
//     );
//     if (classTitle) {
//       title = stripTags_(htmlDecode_(classTitle)).replace(/\s+/g, " ").trim();
//     }
//   }

//   // 画像URL（リンクに最も近いimg要素）
//   let imageUrl = "";
//   const imgMatches = block.match(/<img[^>]+src="([^"]+)"[^>]*>/gi);
//   if (imgMatches && imgMatches.length > 0) {
//     // 最初に見つかった画像を使用
//     const imgMatch = imgMatches[0].match(/src="([^"]+)"/i);
//     if (imgMatch) {
//       imageUrl = htmlDecode_(imgMatch[1]);
//     }
//   }

//   // 価格（円を含む数値、カンマ付き）- 複数の価格から最小値を選択
//   let price = "";
//   let prices = [];

//   // 価格検出の前に、送料無料ラインの説明文を除去
//   let priceSearchBlock = block;
//   // 送料無料ラインの説明文を除去（3980を誤検出しないため）
//   priceSearchBlock = priceSearchBlock.replace(/送料無料ラインを[^。]+。/g, "");
//   priceSearchBlock = priceSearchBlock.replace(/3,980円以下に設定[^。]+。/g, "");

//   // 商品価格エリアを特定してから価格を抽出
//   const priceAreaMatches = [];

//   // 価格エリアのパターン（優先度順、楽天の様々な価格表示に対応）
//   const priceAreaPatterns = [
//     // 価格専用のspan/div（楽天でよく使われる）
//     /<(?:span|div)[^>]*class="[^"]*(?:price|item-price|important)[^"]*"[^>]*>([^<]*[0-9][^<]*)</gi,
//     // strongタグ内の価格
//     /<strong[^>]*>([0-9,]+)\s*円<\/strong>/gi,
//     // 価格リンク内のテキスト（価格へのリンク）
//     /<a[^>]*(?:class="[^"]*price[^"]*"|href="[^"]*price[^"]*")[^>]*>([^<]+)</gi,
//     // data-price属性（楽天APIデータ）
//     /data-price="([0-9]+)"/gi,
//     // content-priceクラス内
//     /<div[^>]*class="[^"]*content-price[^"]*"[^>]*>[\s\S]*?([0-9,]+)\s*円/gi,
//     // importantクラス内の価格（楽天の重要情報表示）
//     /<(?:div|span)[^>]*class="[^"]*important[^"]*"[^>]*>[\s\S]*?([0-9,]+)\s*円/gi,
//     // 価格を示すテキスト（「円」の前の数値）
//     />([0-9,]+)\s*円<\/(?:span|div|p|strong|b)>/gi,
//     // 明示的な価格表記
//     /(?:価格|商品価格|販売価格)[\s:：]*([0-9,]+)円/gi,
//     // 商品リンク直後の価格（距離を短縮）
//     /item\.rakuten\.co\.jp[^>]*>[\s\S]{0,500}?([0-9,]+)\s*円/gi,
//   ];

//   // 各パターンで価格エリアを検索
//   for (let pattern of priceAreaPatterns) {
//     let match;
//     while ((match = pattern.exec(priceSearchBlock)) !== null) {
//       priceAreaMatches.push(match[1]);
//     }
//   }

//   // 価格エリアから数値を抽出
//   for (let areaText of priceAreaMatches) {
//     const priceMatch = areaText.match(/([0-9]{1,3}(?:,[0-9]{3})*)/);
//     if (priceMatch) {
//       const normalized = normalizeNumber_(priceMatch[1]);
//       if (normalized && normalized !== "0") {
//         const numPrice = parseInt(normalized);
//         if (!isNaN(numPrice) && numPrice >= 100) {
//           // 100円以上を有効な価格とする
//           prices.push(numPrice);
//           if (index < 3) {
//             console.log("価格エリアから発見:", priceMatch[1], "→", numPrice);
//           }
//         }
//       }
//     }
//   }

//   // 価格エリアで見つからない場合は、通常のパターンで検索（フォールバック）
//   if (prices.length === 0) {
//     const pricePatterns = [
//       // 基本パターン：数値+円（送料や送料無料ラインを除外）
//       /(?<!送料|込み)([0-9]{1,3}(?:,[0-9]{3})*)\s*円(?!以下|込み)/g,
//       // 価格の前に¥記号がある場合
//       /¥([0-9]{1,3}(?:,[0-9]{3})*)/g,
//     ];

//     // 全ての価格パターンで価格を収集
//     for (let pattern of pricePatterns) {
//       let match;
//       const globalPattern = new RegExp(pattern.source, "gi");
//       while ((match = globalPattern.exec(priceSearchBlock)) !== null) {
//         const rawPrice = match[1];
//         const normalized = normalizeNumber_(rawPrice);
//         if (normalized && normalized !== "0") {
//           const numPrice = parseInt(normalized);
//           if (!isNaN(numPrice) && numPrice >= 100) {
//             prices.push(numPrice);
//             if (index < 3) {
//               console.log("パターンから発見:", rawPrice, "→", numPrice);
//             }
//           }
//         }
//       }
//     }
//   }

//   // 送料込みフラグを確認（送料込みの価格は除外の候補）
//   const hasFreeShipping =
//     block.indexOf("送料無料") > -1 || block.indexOf("送料込") > -1;

//   if (prices.length > 0) {
//     // 重複を除去してソート
//     prices = [...new Set(prices)].sort((a, b) => a - b);

//     if (index < 3) {
//       console.log("見つかった価格リスト:", prices);
//     }

//     // 価格選択ロジック（改善版）
//     if (prices.length === 1) {
//       // 価格が1つだけの場合はそれを使用
//       price = String(prices[0]);
//     } else if (prices.length > 1) {
//       // 複数の価格がある場合の選択ロジック

//       // Step 1: 3980円（送料無料ライン）を除外
//       let candidates = prices.filter((p) => p !== 3980);

//       if (candidates.length === 0) {
//         // 全て3980の場合は仕方なく使用
//         candidates = prices;
//       }

//       // Step 2: 送料と思われる価格を識別して除外
//       // 一般的な送料: 500-1000円
//       const shippingPrices = [
//         500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000,
//       ];
//       const nonShippingPrices = candidates.filter(
//         (p) => !shippingPrices.includes(p)
//       );

//       if (nonShippingPrices.length > 0) {
//         // 送料以外の価格がある場合
//         // 1000円以上で最も小さい価格を選択（送料込み価格を避ける）
//         const over1000 = nonShippingPrices.filter((p) => p >= 1000);
//         if (over1000.length > 0) {
//           price = String(Math.min(...over1000));
//         } else {
//           // 1000円未満しかない場合は最大値を選択
//           price = String(Math.max(...nonShippingPrices));
//         }
//       } else {
//         // 全て送料と思われる価格の場合
//         // 1000円以上の価格があればその最小値を選択
//         const reasonablePrices = candidates.filter((p) => p >= 1000);
//         if (reasonablePrices.length > 0) {
//           price = String(Math.min(...reasonablePrices));
//         } else {
//           // 全て1000円未満の場合は最大値を使用
//           price = String(Math.max(...candidates));
//         }
//       }

//       // デバッグ情報
//       if (index < 3) {
//         console.log(
//           "価格選択結果:",
//           price,
//           "（候補:",
//           candidates.join(","),
//           "）"
//         );
//       }
//     }
//   }

//   // 価格が見つからない場合のフォールバック
//   if (!price) {
//     // より広範囲で価格を探す（最初に見つかった数値+円）
//     const fallbackMatch = block.match(/([0-9,]+)\s*円/);
//     if (fallbackMatch) {
//       const rawPrice = fallbackMatch[1];
//       const normalized = normalizeNumber_(rawPrice);
//       // 数値化できた場合は数値、できなかった場合は元のテキストを保持
//       price = normalized || rawPrice;
//     }
//   }

//   // デバッグ：価格抽出結果
//   if (index < 3) {
//     console.log("抽出された価格:", price || "価格が見つかりません");
//   }

//   // ショップ名の抽出（複数のパターンで試行）
//   let shop = "";

//   // パターン1: 楽天の一般的なショップ名表示パターン（改善版）
//   const shopPatterns = [
//     // 楽天市場店を含むリンク（最も確実）
//     /<a[^>]*>([^<]*楽天市場店)<\/a>/,
//     // 【】を含むショップ名（銀座パリスなど）
//     /<a[^>]*>(【[^】]+】[^<]*)<\/a>/,
//     // 単純なリンクで「店」を含む
//     /<a[^>]*>([^<]*店[^<]*)<\/a>/,
//     // merchantクラスを含む要素
//     /<(?:a|span|div)[^>]+class="[^"]*merchant[^"]*"[^>]*>([^<]+)<\/(?:a|span|div)>/i,
//     // 送料の前のリンク（シンプル版）
//     /<a[^>]*>([^<]+)<\/a>[^<]*送料/,
//     // content-merchant内のテキスト
//     /<div[^>]*class="[^"]*content-merchant[^"]*"[^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>/i,
//     // shopやstoreを含むhref属性のリンク
//     /<a[^>]*href="[^"]*\/(?:shop|store)\/[^"]*"[^>]*>([^<]+)<\/a>/,
//     // 価格の後のリンク（広範囲）
//     /円[^<]*<\/[^>]+>[^<]*<a[^>]*>([^<]+)<\/a>/,
//     // dui-linkbox内のショップリンク
//     /<(?:a|span)[^>]*class="[^"]*dui-linkbox[^"]*"[^>]*>([^<]+)<\/(?:a|span)>/i,
//     // 39ショップの前のリンク
//     /<a[^>]*>([^<]+)<\/a>[^<]*39ショップ/,
//     // ポイントの前のリンク
//     /<a[^>]*>([^<]+)<\/a>[^<]*ポイント/,
//     // merchantやshop-nameクラス
//     /<[^>]+class="[^"]*(?:merchant-name|shop-name|store-name)[^"]*"[^>]*>([^<]+)<\/[^>]+>/i,
//   ];

//   // 各パターンを試行
//   for (let i = 0; i < shopPatterns.length; i++) {
//     const pattern = shopPatterns[i];
//     const shopMatch = block.match(pattern);
//     if (shopMatch && shopMatch[1]) {
//       let candidate = shopMatch[1].trim();

//       // HTMLエンティティをデコード
//       candidate = htmlDecode_(candidate);

//       if (index < 3) {
//         console.log("パターン" + i + "でマッチ:", candidate);
//       }

//       // ショップ名の妥当性チェック（緩和版）
//       if (
//         candidate &&
//         candidate.length > 1 && // 1文字は除外
//         candidate.length < 100 && // ショップ名は100文字以内
//         candidate !== '">' && // ">のみは除外
//         !candidate.match(/^["><!-]+$/) && // HTMLタグの断片のみは除外
//         !candidate.match(/^[\d\s,]+$/) && // 数字、空白、カンマのみは除外
//         !candidate.startsWith('">') && // ">で始まるものは除外
//         !candidate.includes("<!--") && // HTMLコメントは除外
//         !candidate.match(
//           /^(送料|ポイント|円|価格|個数|在庫|税込|税別|込み|無料|別|もっと見る|詳細を見る|レビュー|カート|かごに入れる)$/i
//         )
//       ) {
//         // 価格関連の単語のみは除外

//         // 日本語、英字、【】、店を含むものはOK
//         if (
//           candidate.match(
//             /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\u3400-\u4DBFa-zA-Z【】]/
//           ) ||
//           candidate.includes("店")
//         ) {
//           shop = candidate;
//           if (index < 3) {
//             console.log("採用されたショップ名:", candidate, "- パターン:", i);
//           }
//           break;
//         }
//       }

//       if (index < 3 && !shop && candidate) {
//         console.log("除外されたショップ名候補:", candidate, "- パターン:", i);
//       }
//     }
//   }

//   // パターン2: URLから店舗名を抽出（フォールバック）
//   if (!shop || shop.length === 0 || shop === '">') {
//     const shopFromUrl = detailUrl.match(/item\.rakuten\.co\.jp\/([^\/]+)\//);
//     if (shopFromUrl) {
//       shop = shopFromUrl[1];
//       if (index < 3) {
//         console.log("URLからショップ名を取得:", shop);
//       }
//     }
//   }

//   // 最終的な妥当性チェック
//   if (shop === '">' || shop.match(/^[">]+$/)) {
//     shop = "";
//   }

//   if (index < 3) {
//     console.log("最終的なショップ名:", shop || "ショップ名が見つかりません");
//     if (!shop || shop === '">') {
//       // ショップ名が見つからない場合、リンクを含むエリアを表示
//       const linkArea = block.match(
//         /<a[^>]+>([^<]{1,50})<\/a>[\s\S]{0,50}(?:送料|ポイント|39ショップ)/g
//       );
//       if (linkArea) {
//         console.log("リンクエリア（最初の3つ）:");
//         linkArea.slice(0, 3).forEach((area, i) => {
//           console.log(i + 1 + ":", area.replace(/\s+/g, " ").substring(0, 100));
//         });
//       }
//     }
//   }

//   // 商品情報を返す
//   return {
//     title: title,
//     detailUrl: detailUrl,
//     imageUrl: imageUrl,
//     date: "", // 楽天には日付情報がないため空
//     rank: "", // 楽天にはランク情報がないため空
//     price: price,
//     shop: shop,
//     source: "楽天",
//   };
// }

// /** 正規表現ユーティリティ */
// function firstMatch_(text, re) {
//   const m = re.exec(text);
//   return m ? (m[1] || "").trim() : "";
// }

// /** エコリング用パーサ */
// function parseEcoringFromHtml_(html) {
//   var items = [];
//   var H = String(html);
//   var baseUrl = "https://www.ecoauc.com";
  
//   console.log("エコリングパーサー - HTMLの長さ:", H.length);
//   console.log("エコリングパーサー - HTMLに'show-case-title-block'が含まれているか:", H.indexOf("show-case-title-block") > -1);
  
//   // 商品ブロックを抽出（col-sm-6 col-md-4 col-lg-3）
//   var blockRe = /<div\s+class="col-sm-6\s+col-md-4\s+col-lg-3"[^>]*>[\s\S]*?(?=<div\s+class="col-sm-6\s+col-md-4\s+col-lg-3"|<\/div>\s*<\/div>\s*<\/div>\s*<\/section>|$)/gi;
//   var blocks = H.match(blockRe);
  
//   if (!blocks || blocks.length === 0) {
//     console.log("エコリング 商品ブロックが見つかりません");
//     return items;
//   }
  
//   console.log("エコリング 商品ブロック数:", blocks.length);
  
//   blocks.forEach(function(block, index) {
//     // 詳細URL（相対パスを絶対パスに変換）
//     var detailUrl = "";
//     var urlMatch = block.match(/<a[^>]+href="(\/client\/market-prices\/view\/[0-9]+)"[^>]*>/i);
//     if (urlMatch) {
//       detailUrl = baseUrl + htmlDecode_(urlMatch[1]);
//     }
    
//     // 画像URL（show-case-img-topクラスの画像を優先）
//     var imageUrl = "";
//     var imgMatch = block.match(/<img[^>]+class="[^"]*show-case-img-top[^"]*"[^>]+src="([^"]+)"/i);
//     if (!imgMatch) {
//       // 属性の順序が異なる場合
//       imgMatch = block.match(/<img[^>]+src="([^"]+)"[^>]+class="[^"]*show-case-img-top[^"]*"/i);
//     }
//     if (imgMatch) {
//       imageUrl = htmlDecode_(imgMatch[1]);
//       // 相対パスの場合は絶対パスに変換
//       if (imageUrl && !imageUrl.startsWith("http")) {
//         imageUrl = baseUrl + imageUrl;
//       }
//     }
    
//     // 商品名（b タグ内）
//     var title = "";
//     var titleMatch = block.match(/<div[^>]+class="[^"]*show-case-title-block[^"]*"[^>]*>[\s\S]*?<b>([^<]+)<\/b>/i);
//     if (titleMatch) {
//       title = htmlDecode_(titleMatch[1]).trim();
//     }
    
//     // 日付（small タグ内）
//     var date = "";
//     var dateMatch = block.match(/<small[^>]+class="[^"]*show-case-daily[^"]*"[^>]*>([^<]+)<\/small>/i);
//     if (dateMatch) {
//       date = htmlDecode_(dateMatch[1]).trim();
//     }
    
//     // ランク（canopy-rank内）
//     var rank = "";
//     var rankMatch = block.match(/<li[^>]+class="[^"]*canopy-2[^"]*"[^>]*>[\s\S]*?<span[^>]+class="[^"]*canopy-rank[^"]*"[^>]*>([^<]+)<\/span>/i);
//     if (rankMatch) {
//       rank = normalizeRank_(htmlDecode_(rankMatch[1]).trim());
//     }
    
//     // 価格（item-text内のshow-value）
//     var price = "";
//     var priceAreaMatch = block.match(/<div[^>]+class="[^"]*item-text[^"]*"[^>]*>[\s\S]*?<span[^>]+class="[^"]*show-value[^"]*"[^>]*>([^<]+)<\/span>/i);
//     if (priceAreaMatch) {
//       var priceText = htmlDecode_(priceAreaMatch[1]).trim();
//       // "&yen;12,400"のような形式から数値を抽出
//       price = normalizeNumber_(priceText.replace(/&yen;/g, "").replace(/[¥￥]/g, ""));
//     }
    
//     // カテゴリ（canopy-3内のcanopy-value）
//     var category = "";
//     var categoryMatch = block.match(/<li[^>]+class="[^"]*canopy-3[^"]*"[^>]*>[\s\S]*?<span[^>]+class="[^"]*canopy-value[^"]*"[^>]*>([^<]+)<\/span>/i);
//     if (categoryMatch) {
//       category = htmlDecode_(categoryMatch[1]).trim();
//     }
    
//     // 形状コード（canopy-7内のcanopy-value）
//     var shapeCode = "";
//     var shapeMatch = block.match(/<li[^>]+class="[^"]*canopy-7[^"]*"[^>]*>[\s\S]*?<span[^>]+class="[^"]*canopy-value[^"]*"[^>]*>([^<]+)<\/span>/i);
//     if (shapeMatch) {
//       shapeCode = htmlDecode_(shapeMatch[1]).trim();
//     }
    
//     // ショップ（market-title内）
//     var shop = "";
//     var shopMatch = block.match(/<span[^>]+class="[^"]*market-title[^"]*"[^>]*>([\s\S]*?)<\/span>/i);
//     if (shopMatch) {
//       var shopHtml = shopMatch[1];
//       // HTMLタグを除去してショップ名を抽出
//       shop = stripTags_(htmlDecode_(shopHtml)).trim();
//       // 【】で囲まれた部分を抽出
//       var shopNameMatch = shop.match(/【([^】]+)】/);
//       if (shopNameMatch) {
//         shop = shopNameMatch[1];
//       }
//     }
    
//     // デバッグ情報（最初の3商品）
//     if (index < 3) {
//       console.log("=== エコリング商品 " + (index + 1) + " ===");
//       console.log("詳細URL:", detailUrl);
//       console.log("画像URL:", imageUrl);
//       console.log("商品名:", title);
//       console.log("日付:", date);
//       console.log("ランク:", rank);
//       console.log("価格:", price);
//       console.log("カテゴリ:", category);
//       console.log("形状コード:", shapeCode);
//       console.log("ショップ:", shop);
//     }
    
//     if (detailUrl || imageUrl || title || date || rank || price) {
//       items.push({
//         title: title,
//         detailUrl: detailUrl,
//         imageUrl: imageUrl,
//         date: date,
//         rank: rank,
//         price: price,
//         shop: shop,
//         source: "エコリング",
//       });
//     }
//   });
  
//   console.log("エコリング パース完了:", items.length + "件");
//   return items;
// }

// /** ヤフオク用パーサ */
// function parseYahooAuctionFromHtml_(html) {
//   var items = [];
//   var H = String(html);

//   console.log("ヤフオクパーサー - HTMLの長さ:", H.length);
//   console.log("ヤフオクパーサー - 'Product'クラスが含まれているか:", H.indexOf('class="Product"') > -1);

//   // 全HTMLから直接商品を探す方法
//   var blocks = [];
  
//   // 各商品の開始位置を見つける（HTMLから直接）
//   var productRegex = /<li\s+class="Product"[^>]*>/gi;
//   var productPositions = [];
//   var match;
//   while ((match = productRegex.exec(H)) !== null) {
//     productPositions.push(match.index);
//   }
  
//   // 各商品ブロックを抽出（最大50個）
//   var maxProducts = Math.min(50, productPositions.length);
//   for (var i = 0; i < maxProducts; i++) {
//     var start = productPositions[i];
//     // 次の商品の開始位置または10000文字先まで
//     var end = (i + 1 < productPositions.length) ? productPositions[i + 1] : start + 10000;
    
//     var block = H.substring(start, end);
//     blocks.push(block);
//   }

//   if (!blocks || blocks.length === 0) {
//     console.log("ヤフオク 商品ブロックが見つかりません");
//     return items;
//   }

//   console.log("ヤフオク 商品ブロック数:", blocks.length);

//   var skippedItems = [];
//   var parsedCount = 0;
  
//   blocks.forEach(function(block, index) {
//     // 商品名とURL（Product__titleLinkから）- 改行を含む場合も考慮
//     var title = "";
//     var detailUrl = "";
    
//     // Product__titleLinkから直接取得（改行対応）
//     var titleLinkMatch = block.match(/<a[^>]*class="Product__titleLink"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i);
//     if (titleLinkMatch) {
//       detailUrl = titleLinkMatch[1];
//       // タイトル内のタグを除去し、改行や余分な空白を除去
//       var titleText = titleLinkMatch[2];
//       titleText = titleText.replace(/<[^>]+>/g, '');  // HTMLタグを除去
//       titleText = titleText.replace(/\s+/g, ' ');     // 改行と余分な空白を除去
//       title = htmlDecode_(titleText).trim();
      
//       // 相対URLの場合は絶対URLに変換
//       if (detailUrl && !detailUrl.startsWith("http")) {
//         detailUrl = "https://auctions.yahoo.co.jp" + detailUrl;
//       }
//     }

//     // 画像URL（Product__imageDataから）
//     var imageUrl = "";
//     var imgMatch = block.match(/<img[^>]*class="Product__imageData"[^>]*src="([^"]+)"/i);
//     if (imgMatch) {
//       imageUrl = htmlDecode_(imgMatch[1]);
//     }

//     // 落札価格（Product__priceValueから）- 改行や空白を考慮
//     var price = "";
//     // 落札ラベルの後の価格を探す
//     var priceAreaMatch = block.match(/<span[^>]*class="Product__label"[^>]*>落札<\/span>[\s\S]*?<span[^>]*class="Product__priceValue[^"]*"[^>]*>([^<]+)<\/span>/i);
//     if (priceAreaMatch) {
//       var priceText = htmlDecode_(priceAreaMatch[1]).trim();
//       // "2,420円"のような形式から数値を抽出
//       price = normalizeNumber_(priceText.replace(/円/g, ""));
//     }

//     // 終了日時（Product__timeから）
//     var date = "";
//     var dateMatch = block.match(/<span[^>]*class="Product__time"[^>]*>([^<]+)<\/span>/i);
//     if (dateMatch) {
//       date = htmlDecode_(dateMatch[1]).trim();
//     }

//     // 出品者評価（Product__ratingValueから）
//     var sellerRating = "";
//     var ratingMatch = block.match(/<span[^>]*class="Product__ratingValue"[^>]*>([^<]+)<\/span>/i);
//     if (ratingMatch) {
//       sellerRating = htmlDecode_(ratingMatch[1]).trim();
//     }

//     // ストア情報（年間ベストストアなど）
//     var shop = "";
//     if (block.indexOf("年間ベストストア") > -1) {
//       shop = "ベストストア";
//     } else if (block.indexOf("ストア") > -1) {
//       shop = "ストア";
//     }

//     // 入札件数（Product__bidから）
//     var bidCount = "";
//     var bidMatch = block.match(/<a[^>]*class="Product__bid"[^>]*>([^<]+)<\/a>/i);
//     if (bidMatch) {
//       bidCount = htmlDecode_(bidMatch[1]).trim();
//     }

//     // デバッグ情報（最初の3商品）
//     if (index < 3) {
//       console.log("=== ヤフオク商品 " + (index + 1) + " ===");
//       console.log("ブロックサイズ:", block.length + "文字");
//       console.log("商品名:", title);
//       console.log("詳細URL:", detailUrl);
//       console.log("画像URL:", imageUrl);
//       console.log("落札価格:", price);
//       console.log("終了日時:", date);
//       console.log("入札件数:", bidCount);
//       console.log("出品者評価:", sellerRating);
//       console.log("ショップ:", shop);
      
//       // より詳細なデバッグ情報
//       if (!title) {
//         console.log("titleLinkMatch結果:", titleLinkMatch ? "見つかった" : "見つからない");
//         if (titleLinkMatch) {
//           console.log("titleLinkMatchの内容:", titleLinkMatch[0]);
//         }
//         // Product__titleLinkが含まれているか確認
//         console.log("ブロックに'Product__titleLink'が含まれているか:", block.indexOf('Product__titleLink') > -1);
//         // ブロックの一部を表示
//         console.log("ブロックの内容（最初の1000文字）:", block.substring(0, 1000));
//       }
//       if (!price) {
//         // 価格エリアを表示
//         var priceArea = block.match(/<div[^>]*class="Product__priceInfo"[^>]*>[\s\S]*?<\/div>/i);
//         if (priceArea) {
//           console.log("価格エリアのHTML（最初の300文字）:", priceArea[0].substring(0, 300));
//         }
//       }
//     }

//     if (title && (detailUrl || price || date)) {
//       items.push({
//         title: title,
//         detailUrl: detailUrl,
//         imageUrl: imageUrl,
//         date: date,
//         rank: "", // ヤフオクにはランク情報なし
//         price: price,
//         shop: shop,
//         source: "ヤフオク",
//       });
//       parsedCount++;
//     } else {
//       // スキップされた商品を記録
//       skippedItems.push({
//         index: index + 1,
//         reason: !title ? "タイトルなし" : "詳細URL・価格・日付すべてなし",
//         title: title || "（取得失敗）",
//         detailUrl: detailUrl || "（なし）",
//         price: price || "（なし）",
//         date: date || "（なし）"
//       });
//     }
//   });

//   console.log("ヤフオク パース完了:", items.length + "件");
  
//   // スキップされた商品の情報を表示
//   if (skippedItems.length > 0) {
//     console.log("\n=== スキップされた商品 ===");
//     console.log("スキップ数:", skippedItems.length + "件");
//     console.log("表示可能な商品数:", parsedCount + "件");
    
//     // 最初の5件のスキップ理由を表示
//     skippedItems.slice(0, 5).forEach(function(skipped) {
//       console.log("\n商品" + skipped.index + ": " + skipped.reason);
//       console.log("  タイトル:", skipped.title);
//       console.log("  URL:", skipped.detailUrl);
//       console.log("  価格:", skipped.price);
//       console.log("  日付:", skipped.date);
//     });
    
//     if (skippedItems.length > 5) {
//       console.log("\n...他" + (skippedItems.length - 5) + "件がスキップされました");
//     }
    
//     // エラーログにも記録
//     logWarning_("parseYahooAuctionFromHtml_", "商品スキップ", 
//       "全" + blocks.length + "件中、" + skippedItems.length + "件がスキップされました（表示: " + items.length + "件）");
//   }
  
//   return items;
// }

// // ヤフオクURLからHTMLをフェッチ（ログイン不要）
// function fetchYahooAuctionHtml_(url) {
//   try {
//     logInfo_("fetchYahooAuctionHtml_", "ヤフオクHTML取得開始", "URL: " + url);
    
//     var options = {
//       method: "get",
//       headers: {
//         "User-Agent":
//           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
//         Accept:
//           "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
//         "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
//       },
//       muteHttpExceptions: true,
//       followRedirects: true,
//     };

//     var response = UrlFetchApp.fetch(url, options);
//     var responseCode = response.getResponseCode();
//     var html = response.getContentText("UTF-8");
    
//     console.log("ヤフオク レスポンスコード:", responseCode);
//     console.log("ヤフオク HTMLサイズ:", html.length);

//     if (responseCode !== 200) {
//       logError_("fetchYahooAuctionHtml_", "ERROR", "HTTPステータスエラー", "ステータスコード: " + responseCode);
//       throw new Error(
//         "ヤフオクからのHTMLの取得に失敗しました。ステータスコード: " +
//           responseCode
//       );
//     }
    
//     // HTMLが有効かチェック
//     if (html.length < 1000) {
//       logWarning_("fetchYahooAuctionHtml_", "HTMLが短すぎます", "HTMLサイズ: " + html.length);
//     }
    
//     logInfo_("fetchYahooAuctionHtml_", "HTML取得成功", "HTMLサイズ: " + html.length);
//     return html;
//   } catch (e) {
//     logError_("fetchYahooAuctionHtml_", "ERROR", "HTMLフェッチエラー", e.message);
//     throw new Error("ヤフオクURLからのHTMLフェッチエラー: " + e.message);
//   }
// }

// /** 楽天詳細ページ用パーサ */
// function parseRakutenDetailFromHtml_(html, detailUrl) {
//   const H = String(html);
//   // タイトル
//   let title = firstMatch_(
//     H,
//     /<span[^>]*class="[^"']*normal_reserve_item_name[^"']*"[^>]*>\s*<b>([\s\S]*?)<\/b>/i
//   );
//   if (!title) {
//     title =
//       firstMatch_(H, /<meta[^>]+itemprop="name"[^>]+content="([^"]+)"/i) ||
//       firstMatch_(H, /<meta[^>]+property="og:title"[^>]+content="([^"]+)"/i);
//     if (title) {
//       // 楽天の店名等を含む場合があるので末尾の店名を緩く除去
//       title = title.replace(/\s*[:：].*?楽天市場店?$/, "");
//     }
//   }

//   // 価格（優先: itemprop="price" → RATパラメータ → 画面表示）
//   let price = firstMatch_(
//     H,
//     /<meta[^>]+itemprop="price"[^>]+content="([0-9]+)"/i
//   );
//   if (!price) {
//     price = firstMatch_(H, /id="ratPrice"[^>]*value="([0-9]+)"/i);
//   }
//   if (!price) {
//     price = firstMatch_(H, /([0-9]{1,3}(?:,[0-9]{3})*)\s*円/i);
//     price = normalizeNumber_(price);
//   }

//   // 画像（優先: og:image）
//   let imageUrl = firstMatch_(
//     H,
//     /<meta[^>]+property="og:image"[^>]+content="([^"]+)"/i
//   );
//   if (!imageUrl) {
//     imageUrl = firstMatch_(
//       H,
//       /<img[^>]+src="(https?:\/\/[^"']+)"[^>]*alt="[^"]*"/i
//     );
//   }

//   // ショップ名はURLから抽出
//   let shop = "";
//   const mShop = (detailUrl || "").match(/item\.rakuten\.co\.jp\/([^\/]+)\//);
//   if (mShop) shop = mShop[1];

//   return {
//     title: title || "",
//     detailUrl: detailUrl || "",
//     imageUrl: imageUrl || "",
//     date: "",
//     rank: "",
//     price: price || "",
//     shop: shop,
//     source: "楽天",
//   };
// }

// /** SBA 詳細ページ用パーサ（詳細に専用構造が無い場合は一覧ブロックの先頭を流用） */
// function parseSbaDetailFromHtml_(html, detailUrl) {
//   const H = String(html);
//   // まず一覧ブロックの先頭要素を既存のロジックで取得
//   let items = parseStarBuyersFromHtml_(H);
//   if (items && items.length > 0) {
//     // 可能なら detailUrl に一致するものを優先
//     if (detailUrl) {
//       const found = items.find(
//         (it) =>
//           (it.detailUrl || "").indexOf(detailUrl) !== -1 ||
//           detailUrl.indexOf(it.detailUrl || "") !== -1
//       );
//       if (found) return found;
//     }
//     return items[0];
//   }

//   // フォールバック：title/price/rank/開催日を汎用検索
//   const title =
//     firstMatch_(H, /<h1[^>]*>([\s\S]*?)<\/h1>/i) ||
//     firstMatch_(
//       H,
//       /<div[^>]+class="[^"']*p-item-detail[^"']*"[\s\S]*?<h1[^>]*>([\s\S]*?)<\/h1>/i
//     ) ||
//     "";
//   const priceRaw = firstMatch_(H, /([0-9]{1,3}(?:,[0-9]{3})*)\s*(?:yen|円)/i);
//   const date = firstMatch_(
//     H,
//     /([0-9]{4}\/[0-9]{2}\/[0-9]{2}\s+[0-9]{2}:[0-9]{2}:[0-9]{2})/
//   );
//   const rank = firstMatch_(H, /data-rank="([^"]+)"/i);
//   const img = firstMatch_(H, /<img[^>]+src="(https?:\/\/[^"']+)"/i);

//   return {
//     title: stripTags_(htmlDecode_(title)).replace(/\s+/g, " ").trim(),
//     detailUrl: detailUrl || "",
//     imageUrl: img || "",
//     date: date || "",
//     rank: normalizeRank_(rank || ""),
//     price: normalizeNumber_(priceRaw || ""),
//     shop: "",
//     source: "SBA",
//   };
// }

// /** C10〜V10 の詳細URLを自動取得して出力 */
// function fillFromDetailUrls(sheetName) {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
//   if (!sh) throw new Error("対象シートが見つかりません");

//   // C10〜V10 を読み取り
//   var urls = sh.getRange(10, COL_C, 1, MAX_OUTPUT_ITEMS).getValues()[0];

//   var items = [];
//   for (var i = 0; i < urls.length; i++) {
//     var u = String(urls[i] || "").trim();
//     if (!u || !/^https?:\/\//i.test(u)) {
//       continue;
//     }

//     try {
//       if (/item\.rakuten\.co\.jp\//i.test(u)) {
//         var htmlR = fetchRakutenHtml_(u);
//         items.push(parseRakutenDetailFromHtml_(htmlR, u));
//       } else if (/starbuyers-global-auction\.com\//i.test(u)) {
//         var htmlS = fetchStarBuyersHtml_(u);
//         items.push(parseSbaDetailFromHtml_(htmlS, u));
//       } else {
//         // 未対応のモールはスキップ
//         console.log("未対応URLのためスキップ:", u);
//       }
//     } catch (e) {
//       console.warn("詳細取得エラー:", u, e && e.message);
//       // 失敗時でもURLは出力
//       items.push({
//         title: "",
//         detailUrl: u,
//         imageUrl: "",
//         date: "",
//         rank: "",
//         price: "",
//         shop: "",
//         source: /rakuten/i.test(u)
//           ? "楽天"
//           : /starbuyers/i.test(u)
//           ? "SBA"
//           : "",
//       });
//     }
//   }

//   if (items.length > MAX_OUTPUT_ITEMS) items = items.slice(0, MAX_OUTPUT_ITEMS);
//   writeItemsToTemplate_B10_V19_(sh, items);
//   // 服サイズの自動設定（既存仕様）
//   handleClothingSizeField_(sh);
// }
// // 5つのサイトのデータを各セクションに配置する関数
// function writeItemsToMultipleSections_(
//   sheet,
//   rakutenItems,
//   sbaItems,
//   aucfanItems,
//   ecoringItems,
//   yahooItems
// ) {
//   // SBA: B12~V25 (行12-25)
//   writeItemsToSpecificSection_(sheet, sbaItems, 12, "SBA");

//   // エコリング: B45~V58 (行45-58)
//   writeItemsToSpecificSection_(sheet, ecoringItems, 45, "エコリング");

//   // 楽天: B78~V91 (行78-91)
//   writeItemsToSpecificSection_(sheet, rakutenItems, 78, "楽天");

//   // オークファン: B111~V124 (行111-124)
//   writeItemsToSpecificSection_(sheet, aucfanItems, 111, "オークファン");

//   // ヤフオク: B144~V157 (行144-157)
//   writeItemsToSpecificSection_(sheet, yahooItems, 144, "ヤフオク");
// }

// // 指定されたセクションにアイテムデータを配置する関数
// function writeItemsToSpecificSection_(sheet, items, startRow, sectionName) {
//   logInfo_("writeItemsToSpecificSection_", "データ配置開始", sectionName + "セクション、" + (items ? items.length : 0) + "件");
  
//   if (!items || items.length === 0) {
//     console.log(sectionName + "のデータがありません");
//     logWarning_("writeItemsToSpecificSection_", "データなし", sectionName + "セクション");
//     return;
//   }

//   // スプレッドシートのレイアウトに合わせた行マッピング
//   var rowMapping = {
//     詳細URL: startRow + 0, // 11行目
//     画像URL: startRow + 1, // 12行目
//     // 画像: startRow + 2,     // 13行目（スキップ）
//     // 計算に入れるか: startRow + 3, // 14行目（スキップ）
//     商品名: startRow + 4, // 15行目
//     日付: startRow + 5, // 16行目
//     ランク: startRow + 6, // 17行目
//     // ランク数値: startRow + 7, // 18行目（スキップ）
//     // 服サイズ: startRow + 8,  // 19行目（別途処理）
//     引用サイト: startRow + 9, // 20行目
//     // カラー: startRow + 10,   // 21行目（スキップ）
//     // 外れ値か？: startRow + 11, // 22行目（スキップ）
//     価格: startRow + 12, // 23行目
//     ショップ: startRow + 13, // 24行目
//   };

//   // 出力件数制限（C〜V）
//   var limitedItems = items.slice(0, MAX_OUTPUT_ITEMS);

//   // データ配列を準備
//   var rowArrays = {
//     詳細URL: limitedItems.map((it) => it.detailUrl || ""),
//     画像URL: limitedItems.map((it) => it.imageUrl || ""),
//     商品名: limitedItems.map((it) => it.title || ""),
//     日付: limitedItems.map((it) => it.date || ""),
//     ランク: limitedItems.map((it) => {
//       var rank = normalizeRank_(it.rank || "");
//       return rank; // 空文字列の場合はそのまま返す（後でフィルタリング）
//     }),
//     引用サイト: limitedItems.map((it) => it.source || ""),
//     価格: limitedItems.map((it) => it.price || ""),
//     ショップ: limitedItems.map((it) => it.shop || ""),
//   };

//   // 各ラベルに対応する行にデータを配置
//   Object.keys(rowMapping).forEach(function (label) {
//     var targetRow = rowMapping[label];
//     var data = rowArrays[label];

//     if (data && data.length > 0) {
//       console.log(
//         sectionName + "の「" + label + "」を行" + targetRow + "に出力:",
//         data.length + "件"
//       );

//       // C列からV列まで（20列）にデータを設定
//       var values = [];
//       for (var i = 0; i < MAX_OUTPUT_ITEMS; i++) {
//         var value = i < data.length ? data[i] : "";

//         // ランクの場合、空文字列は入力しない（データ入力規則違反を防ぐ）
//         if (label === "ランク" && value === "") {
//           values.push(""); // 空文字列のまま
//         } else {
//           values.push(value);
//         }
//       }

//       // 該当行のC列以降をクリア
//       clearRowFromColumnC_(sheet, targetRow);

//       // データを設定（ランクの場合は空でないデータがある場合のみ）
//       var hasValidData = false;
//       if (label === "ランク") {
//         hasValidData = values.some(
//           (v) => v !== "" && v !== null && v !== undefined
//         );
//       } else {
//         hasValidData = values.some((v) => v !== "");
//       }

//       if (hasValidData) {
//         sheet
//           .getRange(targetRow, COL_C, 1, MAX_OUTPUT_ITEMS)
//           .setValues([values]);

//         // 折返し抑止
//         if (APPLY_CLIP_WRAP) {
//           var outLen = Math.min(limitedItems.length, MAX_OUTPUT_COLS);
//           if (outLen > 0) {
//             sheet
//               .getRange(targetRow, COL_C, 1, outLen)
//               .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
//           }
//         }
//       } else if (label === "ランク") {
//         console.log(sectionName + "のランクデータは空のためスキップしました");
//       }
//     }
//   });

//   console.log(
//     sectionName +
//       "セクション（行" +
//       startRow +
//       "～" +
//       (startRow + 12) +
//       "）にデータを配置しました"
//   );
//   logInfo_("writeItemsToSpecificSection_", "データ配置完了", sectionName + "セクション、" + limitedItems.length + "件を配置");
// }

// // 従来の関数（後方互換性のため残す）
// function writeItemsToTemplate_B10_V19_(sheet, items) {
//   var labels = [
//     "商品名",
//     "詳細URL",
//     "画像URL",
//     "日付",
//     "ランク",
//     "価格",
//     "ショップ",
//     "引用サイト",
//   ];
//   var bHeadings = readColumnBHeadings_(sheet);

//   // まず通常整形
//   var rowArrays = {
//     商品名: items.map((it) => it.title || ""),
//     詳細URL: items.map((it) => it.detailUrl || ""),
//     画像URL: items.map((it) => it.imageUrl || ""),
//     日付: items.map((it) => it.date || ""),
//     ランク: items.map((it) => it.rank || ""),
//     価格: items.map((it) => it.price || ""),
//     ショップ: items.map((it) => it.shop || ""),
//     引用サイト: items.map((it) => it.source || ""),
//   };

//   // "ランク"は最終チェックして許可外は '' にする
//   if (rowArrays["ランク"]) {
//     rowArrays["ランク"] = rowArrays["ランク"].map(normalizeRank_);
//   }

//   labels.forEach(function (label) {
//     var rowIndex = bHeadings.indexOf(label);
//     if (rowIndex !== -1) {
//       console.log("ラベル「" + label + "」を行" + (rowIndex + 1) + "に出力");
//       console.log("データ数:", rowArrays[label].length);
//       if (label === "価格" && rowArrays[label].length > 0) {
//         console.log("価格データの最初の5件:", rowArrays[label].slice(0, 5));
//       }
//       writeRowByLabel_(sheet, bHeadings, label, rowArrays[label]);
//       // （前回の対策）折返し抑止などはそのまま
//       if (APPLY_CLIP_WRAP) {
//         var row = rowIndex + 1;
//         var outLen = Math.min((items || []).length, MAX_OUTPUT_COLS);
//         if (outLen > 0) {
//           sheet
//             .getRange(row, 3, 1, outLen)
//             .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
//         }
//       }
//     } else {
//       console.log("ラベル「" + label + "」がB列に見つかりません");
//     }
//   });

//   if (DO_AUTO_RESIZE && items.length > 0) {
//     sheet.autoResizeColumns(
//       3,
//       Math.max(1, Math.min(items.length, MAX_OUTPUT_COLS))
//     );
//   }
// }

// // 全角→半角（英数のみ）
// function toHalfWidthAlphaNum_(s) {
//   return String(s || "").replace(/[！-～]/g, (ch) =>
//     String.fromCharCode(ch.charCodeAt(0) - 0xfee0)
//   );
// }

// // 許可ランクへ正規化：S, SA, A, AB, B, BC, C 以外は '' にする
// function normalizeRank_(rank) {
//   let r = toHalfWidthAlphaNum_(rank).toUpperCase().replace(/\s+/g, "");
//   // よくある表記ゆれ
//   const map = {
//     Ｓ: "S",
//     ＳＡ: "SA",
//     Ａ: "A",
//     ＡＢ: "AB",
//     Ｂ: "B",
//     ＢＣ: "BC",
//     Ｃ: "C",
//   };
//   if (map[rank]) r = map[rank]; // 全角キーへの保険
//   // 2文字以上の全角が混在していたケースも救済
//   r = r
//     .replace("Ａ", "A")
//     .replace("Ｂ", "B")
//     .replace("Ｃ", "C")
//     .replace("Ｓ", "S");
//   const ALLOWED = new Set(["S", "SA", "A", "AB", "B", "BC", "C"]);
//   return ALLOWED.has(r) ? r : "";
// }

// // 各セクションの服サイズを自動設定する関数
// function handleClothingSizeFieldForAllSections_(sheet) {
//   // D2セルの値を取得（査定ジャンル）
//   var categoryValue = sheet.getRange("D2").getValue();
//   if (!categoryValue) {
//     console.log("D2セルが空のため、服サイズ処理をスキップ");
//     return;
//   }

//   var categoryStr = String(categoryValue);
//   console.log("D2セルの値:", categoryStr);

//   // 服サイズの値を決定
//   var valueToSet = "";
//   if (categoryStr.indexOf("服") !== -1) {
//     // 「服」が含まれている場合は空白（ただしデータ入力規則エラーを避けるため、実際には何も入力しない）
//     valueToSet = "";
//     console.log("D2に「服」が含まれているため、服サイズを空白に設定");
//   } else {
//     // 「服」が含まれていない場合は「なし」を設定
//     valueToSet = "なし";
//     console.log("D2に「服」が含まれていないため、服サイズを「なし」に設定");
//   }

//   // 各セクションの服サイズ行を設定（実際のレイアウトに合わせて修正）
//   var clothingSizeRows = [
//     20, // SBAセクション（B12~V25の中の服サイズ行：12+8=20）
//     53, // エコリングセクション（B45~V58の中の服サイズ行：45+8=53）
//     86, // 楽天セクション（B78~V91の中の服サイズ行：78+8=86）
//     118, // オークファンセクション（B110~V123の中の服サイズ行：110+8=118）
//     151, // ヤフオクセクション（B143~V156の中の服サイズ行：143+8=151）
//   ];

//   // C列からV列まで（20列）同じ値を設定する配列を準備
//   var values = [];
//   for (var i = 0; i < MAX_OUTPUT_ITEMS; i++) {
//     values.push(valueToSet);
//   }

//   // 各セクションの服サイズ行に値を設定
//   // 服の場合（空白の場合）は何も設定しない
//   if (valueToSet !== "") {
//     clothingSizeRows.forEach(function (row) {
//       try {
//         // 151行目の場合は特別にログを出力
//         if (row === 151) {
//           console.log("=== 151行目（ヤフオク服サイズ）の設定開始 ===");
//           console.log("設定する値:", valueToSet);
//           console.log("設定範囲: C151〜V151");
//         }
        
//         sheet.getRange(row, COL_C, 1, MAX_OUTPUT_ITEMS).setValues([values]);
//         console.log("行" + row + "の服サイズに「" + valueToSet + "」を設定しました");
        
//         if (row === 151) {
//           // 設定後の値を確認
//           var checkValue = sheet.getRange(151, COL_C).getValue();
//           console.log("151行目C列の設定後の値:", checkValue);
//         }
//       } catch (e) {
//         console.error("行" + row + "の設定でエラー:", e.message);
//       }
//     });
//   } else {
//     console.log("「服」カテゴリのため、服サイズフィールドはスキップしました");
//     console.log("スキップされた行:", clothingSizeRows.join(", "));
//   }
// }

// // D2の査定ジャンルに基づいてB16（服サイズ）を自動設定（従来の関数、後方互換性のため残す）
// // D2に「服」が含まれていれば空白、含まれていなければ「なし」を設定
// function handleClothingSizeField_(sheet) {
//   // D2セルの値を取得（査定ジャンル）
//   var categoryValue = sheet.getRange("D2").getValue();
//   if (!categoryValue) return;

//   var categoryStr = String(categoryValue);

//   // B列の見出しを取得
//   var bHeadings = readColumnBHeadings_(sheet);

//   // 「服サイズ」の行を探す（通常はB16）
//   var clothingSizeRow = -1;
//   for (var i = 0; i < bHeadings.length; i++) {
//     if (bHeadings[i] === "服サイズ") {
//       clothingSizeRow = i + 1;
//       break;
//     }
//   }

//   if (clothingSizeRow === -1) return; // 服サイズ行が見つからない

//   // C〜V列（20列）に値を設定
//   var valueToSet = "";
//   if (categoryStr.indexOf("服") !== -1) {
//     // 「服」が含まれている場合は空白
//     valueToSet = "";
//     console.log("D2に「服」が含まれているため、服サイズを空白に設定");
//   } else {
//     // 「服」が含まれていない場合は「なし」を設定
//     valueToSet = "なし";
//     console.log("D2に「服」が含まれていないため、服サイズを「なし」に設定");
//   }

//   // C列からV列まで（20列）同じ値を設定
//   // 服の場合（空白の場合）は何も設定しない
//   if (valueToSet !== "") {
//     var values = [];
//     for (var i = 0; i < MAX_OUTPUT_ITEMS; i++) {
//       values.push(valueToSet);
//     }
//     sheet
//       .getRange(clothingSizeRow, COL_C, 1, MAX_OUTPUT_ITEMS)
//       .setValues([values]);
//     console.log("服サイズ行（B" + clothingSizeRow + "）をC〜V列に設定しました");
//   } else {
//     console.log("「服」カテゴリのため、服サイズフィールドはスキップしました");
//   }
// }

// /** ============== URL自動生成機能 ============== **/
// // G6、G7、G8を読み取って、選択されたサイトの検索URLを自動生成
// // G6: サイト選択（楽天、エコリング、SBA、オークファン、ヤフオク、全て）
// // G7: 検索キーワード（例：「プラダ テスート サフィアーノ」）
// // G8: 期間（SBAのみ、日付またはYYYY-MM-DD形式の文字列、例：「2025-02-01」）
// function generateSearchUrl() {
//   try {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var sheet = ss.getActiveSheet();

//     // G6: サイト選択、G7: フリーワード、G8: 期間
//     var siteSelection = sheet.getRange("G6").getValue();
//     var keyword = sheet.getRange("G7").getValue();
//     var period = sheet.getRange("G8").getValue();

//     logInfo_("generateSearchUrl", "URL生成開始", "サイト: " + siteSelection + ", キーワード: " + keyword);

//     if (!keyword) {
//       console.log("キーワードが入力されていません");
//       logWarning_("generateSearchUrl", "キーワードが入力されていません", "");
//       return;
//     }

//     if (!siteSelection) {
//       console.log("サイトが選択されていません");
//       logWarning_("generateSearchUrl", "サイトが選択されていません", "");
//       return;
//     }

//     // D2からジャンルを取得
//     var genre = sheet.getRange("D2").getValue();
//     console.log("取得したサイト選択:", siteSelection);
//     console.log("取得したジャンル:", genre);
//     console.log("取得したキーワード:", keyword);
//     console.log("取得した期間:", period);

//     // サイト選択に応じてURLを生成・設定
//     if (siteSelection === "全て") {
//       generateAllSiteUrls_(sheet, genre, keyword, period);
//     } else if (siteSelection === "楽天") {
//       generateRakutenUrl_(sheet, genre, keyword);
//     } else if (siteSelection === "SBA") {
//       generateSbaUrl_(sheet, genre, keyword, period);
//     } else if (siteSelection === "エコリング") {
//       generateEcoringUrl_(sheet, genre, keyword, period);
//     } else if (siteSelection === "オークファン") {
//       generateAucfanUrl_(sheet, genre, keyword, period);
//     } else if (siteSelection === "ヤフオク") {
//       generateYahooAuctionUrl_(sheet, genre, keyword, period);
//     } else {
//       console.log("未対応のサイトが選択されました:", siteSelection);
//       logWarning_("generateSearchUrl", "未対応のサイト選択", siteSelection);
//     }
    
//     logInfo_("generateSearchUrl", "URL生成完了", "サイト: " + siteSelection);
//   } catch (e) {
//     console.error("=== generateSearchUrl エラー詳細 ===");
//     console.error("エラーメッセージ:", e.message);
//     console.error("エラー名:", e.name);
//     console.error("スタックトレース:", e.stack);
//     console.error("選択されたサイト:", siteSelection);
//     console.error("キーワード:", keyword);
//     console.error("期間:", period);
    
//     // エラーログシートに詳細を記録
//     logError_("generateSearchUrl", "ERROR", 
//       "URL生成でエラー発生 - サイト: " + siteSelection + ", キーワード: " + keyword, 
//       "エラー: " + e.message + "\n" + 
//       "スタックトレース: " + e.stack);
    
//     // ユーザーに分かりやすいエラーメッセージを表示
//     SpreadsheetApp.getUi().alert(
//       "URL生成エラー",
//       "エラーが発生しました:\n\n" + 
//       e.message + "\n\n" +
//       "選択されたサイト: " + siteSelection + "\n" +
//       "詳細はエラーログシートを確認してください。",
//       SpreadsheetApp.getUi().ButtonSet.OK
//     );
    
//     throw e;
//   }
// }

// /** ============== サイト別URL生成関数 ============== **/

// // 全サイトのURLを生成
// function generateAllSiteUrls_(sheet, genre, keyword, period) {
//   var errors = [];
  
//   // 楽天
//   try {
//     generateRakutenUrl_(sheet, genre, keyword);
//   } catch (e) {
//     console.error("楽天URL生成エラー:", e.message);
//     errors.push("楽天: " + e.message);
//     logError_("generateAllSiteUrls_", "楽天URL生成エラー", e.message, e.stack);
//   }
  
//   // SBA
//   try {
//     generateSbaUrl_(sheet, genre, keyword, period);
//   } catch (e) {
//     console.error("SBA URL生成エラー:", e.message);
//     errors.push("SBA: " + e.message);
//     logError_("generateAllSiteUrls_", "SBA URL生成エラー", e.message, e.stack);
//   }
  
//   // エコリング
//   try {
//     generateEcoringUrl_(sheet, genre, keyword, period);
//   } catch (e) {
//     console.error("エコリング URL生成エラー:", e.message);
//     errors.push("エコリング: " + e.message);
//     logError_("generateAllSiteUrls_", "エコリング URL生成エラー", e.message, e.stack);
//   }
  
//   // オークファン
//   try {
//     generateAucfanUrl_(sheet, genre, keyword, period);
//   } catch (e) {
//     console.error("オークファン URL生成エラー:", e.message);
//     errors.push("オークファン: " + e.message);
//     logError_("generateAllSiteUrls_", "オークファン URL生成エラー", e.message, e.stack);
//   }
  
//   // ヤフオク
//   try {
//     generateYahooAuctionUrl_(sheet, genre, keyword, period);
//   } catch (e) {
//     console.error("ヤフオク URL生成エラー:", e.message);
//     errors.push("ヤフオク: " + e.message);
//     logError_("generateAllSiteUrls_", "ヤフオク URL生成エラー", e.message, e.stack);
//   }
  
//   if (errors.length > 0) {
//     logWarning_("generateAllSiteUrls_", "一部のサイトでエラー発生", errors.join("\n"));
//     throw new Error("URL生成で以下のエラーが発生しました:\n" + errors.join("\n"));
//   }
  
//   console.log("全サイトのURLを生成しました");
//   console.log("オークファンは手動でHTMLを173行目以降に貼り付けてください");
// }

// // 楽天のURL生成
// function generateRakutenUrl_(sheet, genre, keyword) {
//   var encodedKeyword = encodeURIComponent(keyword);
//   var rakutenCategoryId = getRakutenCategoryId_(genre);
//   console.log("楽天カテゴリID:", rakutenCategoryId);

//   var rakutenUrl = "";
//   if (rakutenCategoryId) {
//     rakutenUrl =
//       "https://search.rakuten.co.jp/search/mall/" +
//       encodedKeyword +
//       "/" +
//       rakutenCategoryId +
//       "/?f=0&s=4";
//   } else {
//     rakutenUrl =
//       "https://search.rakuten.co.jp/search/mall/" +
//       encodedKeyword +
//       "/?f=0&s=4";
//   }

//   sheet.getRange(77, COL_B).setValue(rakutenUrl);
//   console.log("B77に楽天の検索URLを設定しました:", rakutenUrl);
// }

// // SBAのURL生成
// function generateSbaUrl_(sheet, genre, keyword, period) {
//   var encodedKeyword = encodeURIComponent(keyword);
//   var sbaCategoryId = getSbaCategoryId_(genre);
//   console.log("SBAカテゴリID:", sbaCategoryId);

//   var categoryParam = sbaCategoryId ? sbaCategoryId : "";
//   var periodParam = "";

//   // 日付をYYYY-MM-DD形式に変換
//   if (period) {
//     if (period instanceof Date) {
//       // Dateオブジェクトの場合
//       var year = period.getFullYear();
//       var month = ("0" + (period.getMonth() + 1)).slice(-2);
//       var day = ("0" + period.getDate()).slice(-2);
//       periodParam = year + "-" + month + "-" + day;
//       console.log("日付をフォーマット:", periodParam);
//     } else {
//       // 文字列の場合（YYYY-MM-DD形式を期待）
//       periodParam = String(period);
//       // もし文字列が日付っぽくない場合は、パースを試みる
//       if (
//         periodParam.indexOf("GMT") > -1 ||
//         periodParam.indexOf("日本標準時") > -1
//       ) {
//         var dateObj = new Date(periodParam);
//         if (!isNaN(dateObj.getTime())) {
//           var year = dateObj.getFullYear();
//           var month = ("0" + (dateObj.getMonth() + 1)).slice(-2);
//           var day = ("0" + dateObj.getDate()).slice(-2);
//           periodParam = year + "-" + month + "-" + day;
//           console.log("日付文字列をフォーマット:", periodParam);
//         }
//       }
//     }
//   }

//   var sbaUrl =
//     "https://www.starbuyers-global-auction.com/market_price?sort=exhibit_date&direction=desc&limit=30" +
//     "&item_category=" +
//     categoryParam +
//     "&keywords=" +
//     encodedKeyword +
//     "&rank=&price_from=&price_to=" +
//     "&exhibit_date_from=" +
//     periodParam +
//     "&exhibit_date_to=&accessory=";

//   sheet.getRange(11, COL_B).setValue(sbaUrl);
//   console.log("B11にSBAの検索URLを設定しました:", sbaUrl);
// }

// // エコリングのURL生成
// function generateEcoringUrl_(sheet, genre, keyword, period) {
//   var encodedKeyword = encodeURIComponent(keyword);
  
//   // ジャンル整理シートからカテゴリIDを取得
//   var categoryIds = getEcoringCategoryIds_(genre);
//   console.log("エコリング カテゴリID:", categoryIds);
  
//   // 基本URLの構築
//   var baseUrl = "https://www.ecoauc.com/client/market-prices";
//   var params = [];
  
//   // 基本パラメータ
//   params.push("limit=50");
//   params.push("sortKey=1");
//   params.push("q=" + encodedKeyword);
//   params.push("low=");
//   params.push("high=");
//   params.push("master_item_brands=");
  
//   // カテゴリパラメータ（複数可）
//   if (categoryIds && categoryIds.length > 0) {
//     categoryIds.forEach(function(id, index) {
//       params.push("master_item_categories%5B" + index + "%5D=" + id);
//     });
//   }
  
//   params.push("master_item_shapes=");
  
//   // 日付パラメータ（SBAと同様にperiodから取得）
//   var startYear, startMonth;
//   var now = new Date();
//   var currentYear = now.getFullYear();
//   var currentMonth = now.getMonth() + 1;
  
//   if (period) {
//     if (period instanceof Date) {
//       startYear = period.getFullYear();
//       startMonth = period.getMonth() + 1;
//       console.log("エコリング 開始日付:", startYear + "年" + startMonth + "月");
//     } else {
//       // 文字列の場合、パースを試みる
//       var dateObj = new Date(period);
//       if (!isNaN(dateObj.getTime())) {
//         startYear = dateObj.getFullYear();
//         startMonth = dateObj.getMonth() + 1;
//         console.log("エコリング 開始日付（文字列から変換）:", startYear + "年" + startMonth + "月");
//       } else {
//         // パースできない場合は現在年の1月から
//         startYear = currentYear;
//         startMonth = 1;
//         console.log("エコリング 日付パース失敗、デフォルト使用:", startYear + "年" + startMonth + "月");
//       }
//     }
//   } else {
//     // periodが指定されていない場合は現在年の1月から
//     startYear = currentYear;
//     startMonth = 1;
//     console.log("エコリング period未指定、デフォルト使用:", startYear + "年" + startMonth + "月");
//   }
  
//   params.push("target_start_year=" + startYear);
//   params.push("target_start_month=" + startMonth);
//   params.push("target_end_year=" + currentYear);
//   params.push("target_end_month=" + currentMonth);
  
//   // その他のパラメータ
//   params.push("master_invoice_setting_id=0");
//   params.push("method=1");
//   params.push("master_item_ranks=");
//   params.push("tab=general");
//   params.push("tableType=grid");
  
//   var ecoringUrl = baseUrl + "?" + params.join("&");
  
//   sheet.getRange(44, COL_B).setValue(ecoringUrl); // B44に設定
//   console.log("B44にエコリングの検索URLを設定しました:", ecoringUrl);
// }

// // ヤフオクのURL生成
// function generateYahooAuctionUrl_(sheet, genre, keyword, period) {
//   var encodedKeyword = encodeURIComponent(keyword);
//   var yahooCategoryId = getYahooCategoryId_(genre);
//   console.log("ヤフオク カテゴリID:", yahooCategoryId);
  
//   logInfo_("generateYahooAuctionUrl_", "ヤフオクURL生成", "キーワード: " + keyword + ", カテゴリID: " + yahooCategoryId);

//   var yahooUrl = "";
//   if (yahooCategoryId) {
//     yahooUrl =
//       "https://auctions.yahoo.co.jp/closedsearch/closedsearch?auccat=" +
//       yahooCategoryId +
//       "&tab_ex=commerce&ei=utf-8&aq=-1&oq=&sc_i=&p=" +
//       encodedKeyword;
//   } else {
//     // カテゴリIDがない場合は全体検索
//     yahooUrl =
//       "https://auctions.yahoo.co.jp/closedsearch/closedsearch?" +
//       "tab_ex=commerce&ei=utf-8&aq=-1&oq=&sc_i=&p=" +
//       encodedKeyword;
//   }

//   sheet.getRange(143, COL_B).setValue(yahooUrl); // B143に設定
//   console.log("B143にヤフオクの検索URLを設定しました:", yahooUrl);
// }

// // 卸先シートからカテゴリIDを取得
// function getCategoryId_(genre, siteType) {
//   if (!genre) return "";

//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var genreSheet = ss.getSheetByName("卸先");

//   if (!genreSheet) {
//     console.log("「卸先」シートが見つかりません");
//     return "";
//   }

//   // A列から最終行を取得
//   var lastRow = genreSheet.getLastRow();
//   if (lastRow < 2) return ""; // ヘッダー行のみの場合

//   // C列（ジャンル）の値を取得
//   var genreRange = genreSheet.getRange(2, 3, lastRow - 1, 1).getValues();

//   // ジャンルが一致する行を探す
//   var targetRow = -1;
//   for (var i = 0; i < genreRange.length; i++) {
//     if (genreRange[i][0] === genre) {
//       targetRow = i + 2; // 実際の行番号（ヘッダー行を考慮）
//       break;
//     }
//   }

//   if (targetRow === -1) {
//     console.log(
//       "ジャンル「" + genre + "」がジャンル整理シートに見つかりません"
//     );
//     return "";
//   }

//   // カテゴリIDを取得（I列:SBA、J列:楽天）
//   var categoryId = "";
//   if (siteType === "SBA") {
//     categoryId = genreSheet.getRange(targetRow, 9).getValue(); // I列
//   } else if (siteType === "楽天") {
//     categoryId = genreSheet.getRange(targetRow, 10).getValue(); // J列
//   }

//   return String(categoryId || "");
// }

// // ジャンル整理シートからエコリングのカテゴリIDを取得（K列、複数可）
// function getEcoringCategoryIds_(genre) {
//   if (!genre) return [];

//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var genreSheet = ss.getSheetByName("ジャンル整理");
  
//   // "ジャンル整理"シートが見つからない場合は"卸先"シートを試す
//   if (!genreSheet) {
//     genreSheet = ss.getSheetByName("卸先");
//   }

//   if (!genreSheet) {
//     console.log("「ジャンル整理」または「卸先」シートが見つかりません");
//     return [];
//   }

//   // A列から最終行を取得
//   var lastRow = genreSheet.getLastRow();
//   if (lastRow < 2) return []; // ヘッダー行のみの場合

//   // D2のジャンルに対応するカテゴリIDを取得する
//   // 現在選択されているジャンルの行を探す
//   var selectedGenre = String(genre);
//   console.log("エコリング 検索対象ジャンル:", selectedGenre);
  
//   // K列（11列目）の値を直接取得（D2の値に基づく）
//   // 複数のカテゴリIDがある場合があるため、配列として返す
//   var categoryIds = [];
  
//   // C列（3列目）でジャンルを検索
//   var genreRange = genreSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  
//   for (var i = 0; i < genreRange.length; i++) {
//     if (genreRange[i][0] === selectedGenre) {
//       var targetRow = i + 2;
//       var categoryId = genreSheet.getRange(targetRow, 11).getValue(); // K列
      
//       if (categoryId && String(categoryId).trim() !== "") {
//         // カンマ区切りで複数のIDがある場合を考慮
//         var ids = String(categoryId).split(",");
//         ids.forEach(function(id) {
//           var trimmedId = id.trim();
//           if (trimmedId) {
//             categoryIds.push(trimmedId);
//           }
//         });
//       }
//       break;
//     }
//   }
  
//   console.log("エコリング カテゴリID取得結果:", categoryIds);
//   return categoryIds;
// }

// // ジャンル整理シートからSBAのカテゴリIDを取得（I列）
// function getSbaCategoryId_(genre) {
//   try {
//     if (!genre) {
//       console.log("getSbaCategoryId_: ジャンルが空です");
//       return "";
//     }
    
//     console.log("getSbaCategoryId_: ジャンル =", genre);
    
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var genreSheet = ss.getSheetByName("ジャンル整理");
    
//     if (!genreSheet) {
//       var errorMsg = "「ジャンル整理」シートが見つかりません";
//       console.error(errorMsg);
//       logError_("getSbaCategoryId_", "シートエラー", errorMsg, "");
//       throw new Error(errorMsg);
//     }
    
//     // A列から最終行を取得
//     var lastRow = genreSheet.getLastRow();
//     if (lastRow < 2) {
//       console.log("データ行がありません（ヘッダー行のみ）");
//       return "";
//     }
    
//     // C列（ジャンル）の値を取得
//     var genreRange = genreSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    
//     // ジャンルが一致する行を探す
//     var targetRow = -1;
//     for (var i = 0; i < genreRange.length; i++) {
//       if (genreRange[i][0] === genre) {
//         targetRow = i + 2; // 実際の行番号（ヘッダー行を考慮）
//         break;
//       }
//     }
  
//     if (targetRow === -1) {
//       console.log("ジャンル「" + genre + "」がジャンル整理シートに見つかりません");
//       return "";
//     }
    
//     // I列（9列目）からカテゴリIDを取得
//     var categoryId = genreSheet.getRange(targetRow, 9).getValue();
//     console.log("I列から取得したSBAカテゴリID:", categoryId);
    
//     return String(categoryId || "");
//   } catch (e) {
//     console.error("getSbaCategoryId_エラー:", e.message);
//     console.error("スタックトレース:", e.stack);
//     logError_("getSbaCategoryId_", "カテゴリID取得エラー", 
//       "ジャンル: " + genre + ", エラー: " + e.message, e.stack);
//     throw e;
//   }
// }

// // ジャンル整理シートから楽天のカテゴリIDを取得（J列）
// function getRakutenCategoryId_(genre) {
//   try {
//     if (!genre) {
//       console.log("getRakutenCategoryId_: ジャンルが空です");
//       return "";
//     }
    
//     console.log("getRakutenCategoryId_: ジャンル =", genre);
    
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var genreSheet = ss.getSheetByName("ジャンル整理");
    
//     if (!genreSheet) {
//       var errorMsg = "「ジャンル整理」シートが見つかりません";
//       console.error(errorMsg);
//       logError_("getRakutenCategoryId_", "シートエラー", errorMsg, "");
//       throw new Error(errorMsg);
//     }
    
//     // A列から最終行を取得
//     var lastRow = genreSheet.getLastRow();
//     if (lastRow < 2) {
//       console.log("データ行がありません（ヘッダー行のみ）");
//       return "";
//     }
    
//     // C列（ジャンル）の値を取得
//     var genreRange = genreSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    
//     // ジャンルが一致する行を探す
//     var targetRow = -1;
//     for (var i = 0; i < genreRange.length; i++) {
//       if (genreRange[i][0] === genre) {
//         targetRow = i + 2; // 実際の行番号（ヘッダー行を考慮）
//         break;
//       }
//     }
  
//     if (targetRow === -1) {
//       console.log("ジャンル「" + genre + "」がジャンル整理シートに見つかりません");
//       return "";
//     }
    
//     // J列（10列目）からカテゴリIDを取得
//     var categoryId = genreSheet.getRange(targetRow, 10).getValue();
//     console.log("J列から取得した楽天カテゴリID:", categoryId);
    
//     return String(categoryId || "");
//   } catch (e) {
//     console.error("getRakutenCategoryId_エラー:", e.message);
//     console.error("スタックトレース:", e.stack);
//     logError_("getRakutenCategoryId_", "カテゴリID取得エラー", 
//       "ジャンル: " + genre + ", エラー: " + e.message, e.stack);
//     throw e;
//   }
// }

// // ジャンル整理シートからヤフオクのカテゴリIDを取得（L列）
// function getYahooCategoryId_(genre) {
//   try {
//     if (!genre) {
//       console.log("getYahooCategoryId_: ジャンルが空です");
//       return "";
//     }
    
//     console.log("getYahooCategoryId_: ジャンル =", genre);
    
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var genreSheet = ss.getSheetByName("ジャンル整理");
    
//     // "ジャンル整理"シートが見つからない場合は"卸先"シートを試す
//     if (!genreSheet) {
//       console.log("「ジャンル整理」シートが見つかりません。「卸先」シートを探します");
//       genreSheet = ss.getSheetByName("卸先");
//     }
    
//     if (!genreSheet) {
//       var errorMsg = "「ジャンル整理」または「卸先」シートが見つかりません";
//       console.error(errorMsg);
//       logError_("getYahooCategoryId_", "シートエラー", errorMsg, "");
//       throw new Error(errorMsg);
//     }
    
//     // A列から最終行を取得
//     var lastRow = genreSheet.getLastRow();
//     if (lastRow < 2) {
//       console.log("データ行がありません（ヘッダー行のみ）");
//       return "";
//     }
    
//     // C列（ジャンル）の値を取得
//     var genreRange = genreSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    
//     // ジャンルが一致する行を探す
//     var targetRow = -1;
//     for (var i = 0; i < genreRange.length; i++) {
//       if (genreRange[i][0] === genre) {
//         targetRow = i + 2; // 実際の行番号（ヘッダー行を考慮）
//         break;
//       }
//     }
  
//     if (targetRow === -1) {
//       console.log("ジャンル「" + genre + "」がジャンル整理シートに見つかりません");
//       return "";
//     }
    
//     // L列（12列目）からカテゴリIDを取得
//     var categoryId = genreSheet.getRange(targetRow, 12).getValue();
//     console.log("L列から取得したカテゴリID:", categoryId);
    
//     return String(categoryId || "");
//   } catch (e) {
//     console.error("getYahooCategoryId_エラー:", e.message);
//     console.error("スタックトレース:", e.stack);
//     logError_("getYahooCategoryId_", "カテゴリID取得エラー", 
//       "ジャンル: " + genre + ", エラー: " + e.message, e.stack);
//     throw e;
//   }
// }

// /** ============== 認証情報管理（推奨） ============== **/
// // スクリプトプロパティに認証情報を保存する（初回設定用）
// function setStarBuyersCredentials() {
//   // この関数を一度実行して認証情報を保存してください
//   var scriptProperties = PropertiesService.getScriptProperties();
//   scriptProperties.setProperty("STARBUYERS_EMAIL", "your_email@example.com");
//   scriptProperties.setProperty("STARBUYERS_PASSWORD", "your_password");
// }

// // スクリプトプロパティから認証情報を取得
// function getStarBuyersCredentials_() {
//   var scriptProperties = PropertiesService.getScriptProperties();
//   var email = scriptProperties.getProperty("STARBUYERS_EMAIL");
//   var password = scriptProperties.getProperty("STARBUYERS_PASSWORD");

//   // スクリプトプロパティに設定されていない場合はハードコードされた値を使用（非推奨）
//   if (!email || !password) {
//     console.warn(
//       "認証情報がスクリプトプロパティに設定されていません。setStarBuyersCredentials()を実行してください。"
//     );
//     return {
//       email: "inui.hur@gmail.com",
//       password: "hur22721",
//     };
//   }

//   return { email: email, password: password };
// }

// /** ============== メイン ============== **/
// // 選択されたサイトのデータを取得して各セクションに配置
// // SBA: B12~V25、エコリング: B45~V58、楽天: B78~V91、オークファン: B110~V123、ヤフオク: B143~V156
// // 使い方：
// //   1. G6でサイトを選択し、generateSearchUrl関数でURLを自動生成
// //   2. オークファンの場合はHTMLを173行目以降に手動で貼り付け
// //   3. この関数を実行すると、選択されたサイトのデータを自動取得して配置
// function fillTemplateFromHtml_B10_V19(sheetName) {
//   try {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     var sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
//     if (!sh) throw new Error("対象シートが見つかりません");

//     // G6からサイト選択を取得
//     var siteSelection = sh.getRange("G6").getValue();
//     console.log("メイン関数 - 選択されたサイト:", siteSelection);
//     logInfo_("fillTemplateFromHtml_B10_V19", "処理開始", "サイト: " + siteSelection);

//     // データ取得前に選択されたサイトのデータをクリア
//     console.log("選択されたサイトのデータをクリアしています...");
//     logInfo_("fillTemplateFromHtml_B10_V19", "データクリア", "サイト: " + siteSelection);
//     clearSelectedSiteData_(sh, siteSelection);
//     console.log("データクリア完了");

//   var rakutenItems = [];
//   var sbaItems = [];
//   var aucfanItems = [];
//   var ecoringItems = [];
//   var yahooItems = [];

//   // サイト選択に応じてデータを取得
//   if (siteSelection === "全て") {
//     // 全サイトのデータを取得
//     sbaItems = getSbaData_(sh);
//     rakutenItems = getRakutenData_(sh);
//     aucfanItems = getAucfanData_(sh);
//     ecoringItems = getEcoringData_(sh);
//     yahooItems = getYahooAuctionData_(sh);
//   } else if (siteSelection === "SBA") {
//     sbaItems = getSbaData_(sh);
//   } else if (siteSelection === "楽天") {
//     rakutenItems = getRakutenData_(sh);
//   } else if (siteSelection === "オークファン") {
//     aucfanItems = getAucfanData_(sh);
//   } else if (siteSelection === "エコリング") {
//     ecoringItems = getEcoringData_(sh);
//   } else if (siteSelection === "ヤフオク") {
//     yahooItems = getYahooAuctionData_(sh);
//   } else {
//     console.log("未対応のサイトが選択されました:", siteSelection);
//     return;
//   }

//   // 各セクションにデータを配置
//   writeItemsToMultipleSections_(
//     sh,
//     rakutenItems,
//     sbaItems,
//     aucfanItems,
//     ecoringItems,
//     yahooItems
//   );

//   // D2の査定ジャンルに基づいて各セクションの服サイズを自動設定
//   try {
//     console.log("=== 服サイズ設定処理開始 ===");
//     handleClothingSizeFieldForAllSections_(sh);
//     console.log("=== 服サイズ設定処理完了 ===");
//   } catch (e) {
//     console.error("服サイズ設定でエラー:", e.message);
//     console.error("スタックトレース:", e.stack);
//     logError_("fillTemplateFromHtml_B10_V19", "服サイズ設定エラー", e.message, e.stack);
//   }
  
//   logInfo_("fillTemplateFromHtml_B10_V19", "処理完了", "全てのデータを配置しました");
//   } catch (e) {
//     console.error("=== fillTemplateFromHtml_B10_V19 エラー詳細 ===");
//     console.error("エラーメッセージ:", e.message);
//     console.error("スタックトレース:", e.stack);
    
//     // エラーログシートに詳細を記録
//     logError_("fillTemplateFromHtml_B10_V19", "ERROR", 
//       "メイン処理でエラー発生", 
//       "エラー: " + e.message + "\n" + 
//       "スタックトレース: " + e.stack);
    
//     // ユーザーに分かりやすいエラーメッセージを表示
//     SpreadsheetApp.getUi().alert(
//       "処理エラー",
//       "エラーが発生しました:\n\n" + 
//       e.message + "\n\n" +
//       "詳細はエラーログシートを確認してください。",
//       SpreadsheetApp.getUi().ButtonSet.OK
//     );
    
//     throw e;
//   }
// }

// /** ============== サイト別データ取得関数 ============== **/

// // SBAのデータを取得
// function getSbaData_(sheet) {
//   var sbaUrl = sheet.getRange(11, COL_B).getValue(); // B11
//   console.log("SBAデータ取得 - URL:", sbaUrl);

//   if (!sbaUrl || !String(sbaUrl).trim()) {
//     console.log("SBA URLが空または無効です");
//     return [];
//   }

//   try {
//     var sbaHtml = "";
//     if (/^https?:\/\//i.test(String(sbaUrl).trim())) {
//       console.log("SBA URLからHTMLを取得中:", String(sbaUrl).trim());
//       sbaHtml = fetchStarBuyersHtml_(String(sbaUrl).trim());
//       console.log("SBA HTMLの長さ:", sbaHtml.length);
//       console.log("SBA HTMLの最初の500文字:", sbaHtml.substring(0, 500));
//     } else {
//       sbaHtml = String(sbaUrl);
//     }
//     var items = parseStarBuyersFromHtml_(sbaHtml);
//     console.log("SBAデータを取得しました:", items.length + "件");
//     return items;
//   } catch (e) {
//     console.warn("SBAデータの取得に失敗:", e.message);
//     return [];
//   }
// }

// // 楽天のデータを取得
// function getRakutenData_(sheet) {
//   var rakutenUrl = sheet.getRange(77, COL_B).getValue(); // B77
//   console.log("楽天データ取得 - URL:", rakutenUrl);

//   if (!rakutenUrl || !String(rakutenUrl).trim()) {
//     console.log("楽天 URLが空または無効です");
//     return [];
//   }

//   try {
//     var rakutenHtml = "";
//     if (/^https?:\/\//i.test(String(rakutenUrl).trim())) {
//       rakutenHtml = fetchRakutenHtml_(String(rakutenUrl).trim());
//     } else {
//       rakutenHtml = String(rakutenUrl);
//     }
//     var items = parseRakutenFromHtml_(rakutenHtml);
//     console.log("楽天データを取得しました:", items.length + "件");
//     return items;
//   } catch (e) {
//     console.warn("楽天データの取得に失敗:", e.message);
//     return [];
//   }
// }

// // オークファンのデータを取得（HTML入力行から）
// function getAucfanData_(sheet) {
//   try {
//     var aucfanHtml = readHtmlFromRow_(sheet);
//     if (aucfanHtml && aucfanHtml.trim()) {
//       // HTMLソースがオークファンかどうかを判定
//       var source = detectSource_(aucfanHtml);
//       if (source === "aucfan") {
//         var items = parseAucfanFromHtml_(aucfanHtml);
//         console.log("オークファンデータを取得しました:", items.length + "件");
//         return items;
//       } else {
//         console.log("HTML入力行のデータはオークファンではありません:", source);
//         return [];
//       }
//     } else {
//       console.log("HTML入力行にデータがありません");
//       return [];
//     }
//   } catch (e) {
//     console.warn("オークファンデータの取得に失敗:", e.message);
//     return [];
//   }
// }

// // エコリングのデータを取得
// function getEcoringData_(sheet) {
//   var ecoringUrl = sheet.getRange(44, COL_B).getValue(); // B44
//   console.log("エコリングデータ取得 - URL:", ecoringUrl);
//   logInfo_("getEcoringData_", "エコリングデータ取得開始", "URL: " + ecoringUrl);

//   if (!ecoringUrl || !String(ecoringUrl).trim()) {
//     console.log("エコリング URLが空または無効です");
//     logWarning_("getEcoringData_", "エコリング URLが空または無効", "URL: " + ecoringUrl);
//     return [];
//   }

//   try {
//     var ecoringHtml = "";
//     if (/^https?:\/\//i.test(String(ecoringUrl).trim())) {
//       console.log("エコリング URLからHTMLを取得中:", String(ecoringUrl).trim());
//       ecoringHtml = fetchEcoringHtml_(String(ecoringUrl).trim());
//       console.log("エコリング HTMLの長さ:", ecoringHtml.length);
//       console.log("エコリング HTMLの最初の500文字:", ecoringHtml.substring(0, 500));
//       logInfo_("getEcoringData_", "HTML取得成功", "HTMLの長さ: " + ecoringHtml.length);
//     } else {
//       ecoringHtml = String(ecoringUrl);
//     }
//     var items = parseEcoringFromHtml_(ecoringHtml);
//     console.log("エコリングデータを取得しました:", items.length + "件");
//     logInfo_("getEcoringData_", "データ取得成功", items.length + "件の商品を取得");
//     return items;
//   } catch (e) {
//     console.warn("エコリングデータの取得に失敗:", e.message);
//     logError_("getEcoringData_", "ERROR", "エコリングデータの取得に失敗", e.message + "\n" + e.stack);
//     return [];
//   }
// }

// // ヤフオクのデータを取得
// function getYahooAuctionData_(sheet) {
//   var yahooUrl = sheet.getRange(143, COL_B).getValue(); // B143
//   console.log("ヤフオクデータ取得 - URL:", yahooUrl);
//   logInfo_("getYahooAuctionData_", "ヤフオクデータ取得開始", "URL: " + yahooUrl);

//   if (!yahooUrl || !String(yahooUrl).trim()) {
//     console.log("ヤフオク URLが空または無効です");
//     logWarning_("getYahooAuctionData_", "ヤフオク URLが空または無効", "URL: " + yahooUrl);
//     return [];
//   }

//   try {
//     var yahooHtml = "";
//     if (/^https?:\/\//i.test(String(yahooUrl).trim())) {
//       console.log("ヤフオク URLからHTMLを取得中:", String(yahooUrl).trim());
//       yahooHtml = fetchYahooAuctionHtml_(String(yahooUrl).trim());
//       console.log("ヤフオク HTMLの長さ:", yahooHtml.length);
//       console.log("ヤフオク HTMLの最初の500文字:", yahooHtml.substring(0, 500));
//       logInfo_("getYahooAuctionData_", "HTML取得成功", "HTMLの長さ: " + yahooHtml.length);
//     } else {
//       yahooHtml = String(yahooUrl);
//     }
//     var items = parseYahooAuctionFromHtml_(yahooHtml);
//     console.log("ヤフオクデータを取得しました:", items.length + "件");
//     logInfo_("getYahooAuctionData_", "データ取得成功", items.length + "件の商品を取得");
//     return items;
//   } catch (e) {
//     console.warn("ヤフオクデータの取得に失敗:", e.message);
//     logError_("getYahooAuctionData_", "ERROR", "ヤフオクデータの取得に失敗", e.message + "\n" + e.stack);
//     return [];
//   }
// }

// /** HTML入力行以降をB列含めて全クリア */
// function clearHtmlFromInputRow(sheetName) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
//   if (!sh) throw new Error("対象シートが見つかりません");
//   const last = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   if (last >= ROW_HTML_INPUT) {
//     sh.getRange(
//       ROW_HTML_INPUT,
//       1,
//       last - ROW_HTML_INPUT + 1,
//       lastCol
//     ).clearContent();
//   }
// }

// /**
//  * 選択されたサイトのデータ行のみをクリア
//  * @param {Sheet} sheet - 対象のシート
//  * @param {string} siteSelection - 選択されたサイト（全て、SBA、エコリング、楽天、オークファン、ヤフオク）
//  */
// function clearSelectedSiteData_(sheet, siteSelection) {
//   var lastCol = sheet.getLastColumn();
//   var widthFromC = Math.max(0, lastCol - (COL_C - 1)); // C〜最終列の幅

//   if (widthFromC <= 0) return;

//   if (siteSelection === "全て") {
//     // 全サイトのデータをクリア
//     clearSectionRows_(sheet, 12, 25, widthFromC, "SBA");
//     clearSectionRows_(sheet, 45, 58, widthFromC, "エコリング");
//     clearSectionRows_(sheet, 78, 91, widthFromC, "楽天");
//     clearSectionRows_(sheet, 111, 124, widthFromC, "オークファン");
//     clearSectionRows_(sheet, 144, 157, widthFromC, "ヤフオク");
//   } else if (siteSelection === "SBA") {
//     clearSectionRows_(sheet, 12, 25, widthFromC, "SBA");
//   } else if (siteSelection === "エコリング") {
//     clearSectionRows_(sheet, 45, 58, widthFromC, "エコリング");
//   } else if (siteSelection === "楽天") {
//     clearSectionRows_(sheet, 78, 91, widthFromC, "楽天");
//   } else if (siteSelection === "オークファン") {
//     clearSectionRows_(sheet, 111, 124, widthFromC, "オークファン");
//   } else if (siteSelection === "ヤフオク") {
//     clearSectionRows_(sheet, 144, 157, widthFromC, "ヤフオク");
//   }

//   console.log(siteSelection + "のデータをクリアしました");
// }

// /**
//  * 5つのセクション（SBA、エコリング、楽天、オークファン、ヤフオク）のデータ行をクリア
//  * SBA: B12~V25、エコリング: B45~V58、楽天: B78~V91、オークファン: B110~V123、ヤフオク: B143~V156
//  * （行ヘッダBは残す）
//  */
// function clearSpecificRows() {
//   var sh = SpreadsheetApp.getActiveSheet();
//   var last = sh.getLastRow();
//   var lastCol = sh.getLastColumn();
//   var widthFromC = Math.max(0, lastCol - (COL_C - 1)); // C〜最終列の幅

//   if (widthFromC <= 0) return;

//   // SBAセクション（B12~V25）をクリア
//   clearSectionRows_(sh, 12, 25, widthFromC, "SBA");

//   // エコリングセクション（B45~V58）をクリア
//   clearSectionRows_(sh, 45, 58, widthFromC, "エコリング");

//   // 楽天セクション（B78~V91）をクリア
//   clearSectionRows_(sh, 78, 91, widthFromC, "楽天");

//   // オークファンセクション（B111~V124）をクリア
//   clearSectionRows_(sh, 111, 124, widthFromC, "オークファン");

//   // ヤフオクセクション（B144~V157）をクリア
//   clearSectionRows_(sh, 144, 157, widthFromC, "ヤフオク");

//   // HTML入力行もクリア
//   clearHtmlFromInputRow("");

//   console.log("全セクションのデータをクリアしました");
// }

// /**
//  * 指定されたセクションの行範囲をクリアするヘルパー関数
//  * 関数が入っている行（14, 23, 47, 56, 80, 89, 112, 121, 145, 154）はスキップする
//  */
// function clearSectionRows_(sheet, startRow, endRow, widthFromC, sectionName) {
//   var last = sheet.getLastRow();

//   if (last >= startRow) {
//     var actualEndRow = Math.min(endRow, last);

//     // 関数が入っている行を定義（保護対象行）
//     var functionRows = [14, 23, 47, 56, 80, 89, 112, 121, 145, 154]; // SBA:14,23 エコリング:47,56 楽天:80,89 オークファン:112,121 ヤフオク:145,154

//     // 各行を個別にクリア（関数行はスキップ）
//     for (var row = startRow; row <= actualEndRow; row++) {
//       if (functionRows.indexOf(row) === -1) {
//         // 関数行でない場合はクリア
//         sheet.getRange(row, COL_C, 1, widthFromC).clearContent();
//       } else {
//         console.log("行" + row + "は関数が入っているためスキップしました");
//       }
//     }

//     // チェックボックス行（計算に入れるか：各セクションの4行目）をfalseに設定
//     var checkboxRow = startRow + 3; // 12+3=15, 45+3=48, 78+3=81, 110+3=113, 143+3=146
//     if (
//       checkboxRow <= actualEndRow &&
//       functionRows.indexOf(checkboxRow) === -1
//     ) {
//       var range = sheet.getRange(checkboxRow, COL_C, 1, widthFromC);
//       var values = range.getValues();
//       for (var i = 0; i < values[0].length; i++) {
//         if (typeof values[0][i] === "boolean") {
//           values[0][i] = false;
//         } else {
//           values[0][i] = ""; // チェックボックス以外は空文字列
//         }
//       }
//       range.setValues(values);
//     }

//     console.log(
//       sectionName +
//         "セクション（行" +
//         startRow +
//         "～" +
//         actualEndRow +
//         "）をクリアしました（関数行は除く）"
//     );
//   }
// }

// /** ==================== オークファンスクレイピング機能 ==================== **/

// /**
//  * オークファンからデータを取得するテスト関数
//  * Cookiesシートに保存されたクッキー情報を使用
//  */
// function testAucfanScraping() {
//   console.log("=== オークファンデータ取得テスト開始 ===");
//   var startTime = new Date().getTime();
  
//   try {
//     // スプレッドシートからクッキーを読み込む
//     var cookies = loadCookiesFromSheet();
//     if (cookies.length === 0) {
//       console.error("❌ クッキーが見つかりません");
//       console.error("=== 対処法 ===");
//       console.error("1. ターミナルを開く");
//       console.error("2. cd /Users/kazuyukijimbo/aicon");
//       console.error("3. python aucfan_sheet_writer.py");
//       console.error("4. ログインとreCAPTCHAを完了");
//       console.error("5. クッキーがスプレッドシートに保存されたら再実行");
      
//       SpreadsheetApp.getActiveSpreadsheet().toast(
//         "クッキーがありません。\n\nターミナルで以下を実行:\npython aucfan_sheet_writer.py", 
//         "❌ 要ログイン", 
//         15
//       );
//       return;
//     }
    
//     console.log("読み込んだクッキー数: " + cookies.length);
    
//     // テスト検索を実行
//     var keyword = "iPhone 15";
//     console.log("検索キーワード: " + keyword);
    
//     // タイムアウトチェック（60秒）
//     var maxExecutionTime = 60000; // 60秒
//     var checkTimeout = function() {
//       var elapsed = new Date().getTime() - startTime;
//       if (elapsed > maxExecutionTime) {
//         throw new Error("実行時間が60秒を超えました。処理を中断します。");
//       }
//     };
    
//     checkTimeout();
//     var results = searchAucfanProducts(keyword, cookies);
//     checkTimeout();
    
//     if (results && results.length > 0) {
//       console.log("検索結果: " + results.length + "件");
      
//       // 結果をシートに保存
//       try {
//         saveAucfanResultsToSheet(results, keyword);
//         SpreadsheetApp.getUi().alert(
//           "✅ 成功: " + results.length + "件の商品を取得しました\n\nAucfanResultsシートを確認してください。"
//         );
//       } catch (saveError) {
//         console.error("結果の保存中にエラー:", saveError.message);
//         SpreadsheetApp.getUi().alert(
//           "⚠️ 警告: 商品取得は成功しましたが、保存中にエラーが発生しました\n\n" + saveError.message
//         );
//       }
//     } else {
//       console.log("検索結果なし");
//       SpreadsheetApp.getUi().alert("商品が見つかりませんでした\n\n以下を確認してください:\n・クッキーが有効か\n・検索キーワードが適切か");
//     }
    
//   } catch (e) {
//     var errorMessage = e.toString();
//     console.error("エラー発生: " + errorMessage);
//     console.error("スタックトレース:", e.stack);
    
//     // エラーの種類に応じたメッセージ
//     if (errorMessage.includes("timeout") || errorMessage.includes("タイムアウト") || errorMessage.includes("60秒")) {
//       SpreadsheetApp.getUi().alert("❌ タイムアウトエラー\n\n処理に時間がかかりすぎました。\nもう一度実行してみてください。");
//     } else if (errorMessage.includes("認証") || errorMessage.includes("ログイン")) {
//       SpreadsheetApp.getUi().alert("❌ 認証エラー\n\nクッキーが無効の可能性があります。\naucfan_sheet_writer.pyで再ログインしてください。");
//     } else {
//       SpreadsheetApp.getUi().alert("❌ エラーが発生しました\n\n" + errorMessage + "\n\n詳細はログを確認してください。");
//     }
//   } finally {
//     var executionTime = (new Date().getTime() - startTime) / 1000;
//     console.log("実行時間: " + executionTime + "秒");
//     console.log("=== オークファンデータ取得テスト終了 ===");
//   }
// }

// /**
//  * Cookiesシートからクッキー情報を読み込む（改善版）
//  * @returns {Array} クッキーオブジェクトの配列
//  */
// function loadCookiesFromSheet() {
//   try {
//     console.log("🔍 Cookiesシートを検索中...");
//     var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cookies');
    
//     if (!sheet) {
//       console.error("\n╔══════════════════════════════════════════════════════╗");
//       console.error("║         ❌ Cookiesシートが見つかりません             ║");
//       console.error("╚══════════════════════════════════════════════════════╝");
//       console.error("\n📝 対処法:");
//       console.error("  1. ターミナルを開く");
//       console.error("  2. 以下のコマンドを実行:");
//       console.error("     ┌─────────────────────────────────────────┐");
//       console.error("     │ cd /Users/kazuyukijimbo/aicon           │");
//       console.error("     │ python aucfan_sheet_writer.py           │");
//       console.error("     └─────────────────────────────────────────┘");
//       console.error("  3. ブラウザでログインを完了");
//       console.error("  4. Cookiesシートが自動作成されます");
//       console.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
      
//       SpreadsheetApp.getActiveSpreadsheet().toast(
//         '❌ Cookiesシートが見つかりません\n\n' +
//         'aucfan_sheet_writer.pyでログインしてください。',
//         '🔴 シートなし',
//         15
//       );
//       return [];
//     }
    
//     console.log("✅ Cookiesシートを検出");
    
//     var lastRow = sheet.getLastRow();
//     if (lastRow <= 1) {
//       console.error("\n╔══════════════════════════════════════════════════════╗");
//       console.error("║         ⚠️ Cookiesシートにデータがありません         ║");
//       console.error("╚══════════════════════════════════════════════════════╝");
//       console.error("\n📝 対処法:");
//       console.error("  1. 以下のコマンドを実行:");
//       console.error("     ┌─────────────────────────────────────────┐");
//       console.error("     │ cd /Users/kazuyukijimbo/aicon           │");
//       console.error("     │ python aucfan_sheet_writer.py           │");
//       console.error("     └─────────────────────────────────────────┘");
//       console.error("  2. ログインとreCAPTCHA解決");
//       console.error("  3. クッキーが自動保存されます");
//       console.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
      
//       SpreadsheetApp.getActiveSpreadsheet().toast(
//         '⚠️ クッキーデータがありません\n\n' +
//         '再度ログインしてください。',
//         '🟡 データなし',
//         15
//       );
//       return [];
//     }
  
//   // A列～I列のデータを取得（ヘッダー行を除く）
//   var dataRange = sheet.getRange(2, 1, lastRow - 1, 9);
//   var data = dataRange.getValues();
  
//   // クッキーオブジェクトに変換
//   var cookies = [];
//   for (var i = 0; i < data.length; i++) {
//     var row = data[i];
//     if (row[2] && row[3]) { // nameとvalueが存在する場合
//       cookies.push({
//         timestamp: row[0],
//         domain: row[1],
//         name: row[2],
//         value: row[3],
//         path: row[4] || '/',
//         expiry: row[5],
//         secure: row[6] === 'True' || row[6] === true,
//         httpOnly: row[7] === 'True' || row[7] === true,
//         sameSite: row[8] || 'None'
//       });
//     }
//   }
  
//   console.log("✅ " + cookies.length + "個のクッキーを読み込みました");
//   return cookies;
  
//   } catch (e) {
//     console.error("❌ クッキー読み込みエラー:", e.message);
//     console.error("スタックトレース:", e.stack);
//     SpreadsheetApp.getActiveSpreadsheet().toast(
//       'クッキーの読み込みに失敗しました: ' + e.message,
//       '❌ エラー',
//       10
//     );
//     return [];
//   }
// }

// /**
//  * クッキーをHTTPヘッダー形式に変換
//  * @param {Array} cookies - クッキーオブジェクトの配列
//  * @returns {string} Cookieヘッダー文字列
//  */
// function formatCookieHeader(cookies) {
//   if (!cookies || cookies.length === 0) {
//     console.warn("⚠️ formatCookieHeader: クッキーが空です");
//     return "";
//   }
  
//   var cookieStrings = [];
//   for (var i = 0; i < cookies.length; i++) {
//     if (cookies[i].name && cookies[i].value) {
//       cookieStrings.push(cookies[i].name + '=' + cookies[i].value);
//     }
//   }
//   return cookieStrings.join('; ');
// }

// /**
//  * クッキーの有効性をチェック（改善版）
//  * @param {Array} cookies - クッキー配列
//  * @returns {Object} チェック結果
//  */
// function validateCookies(cookies) {
//   var result = {
//     isValid: false,
//     hasSessionCookie: false,
//     expiredCookies: [],
//     validCookies: [],
//     totalCookies: 0,
//     sessionInfo: null,
//     message: "",
//     detailedMessage: []
//   };
  
//   if (!cookies || cookies.length === 0) {
//     result.message = "❌ クッキーが見つかりません";
//     result.detailedMessage.push("対処法: aucfan_sheet_writer.pyでログインしてください");
//     return result;
//   }
  
//   result.totalCookies = cookies.length;
//   var now = new Date().getTime() / 1000; // 秒単位のタイムスタンプ
//   var sessionCookies = ['PHPSESSID', 'session', 'laravel_session', 'ci_session', '_session_id'];
  
//   for (var i = 0; i < cookies.length; i++) {
//     var cookie = cookies[i];
    
//     // セッションクッキーチェック
//     if (sessionCookies.indexOf(cookie.name) > -1 || 
//         cookie.name.toLowerCase().indexOf('sess') > -1) {
//       result.hasSessionCookie = true;
      
//       // セッションクッキーの詳細情報を保存
//       result.sessionInfo = {
//         name: cookie.name,
//         domain: cookie.domain,
//         path: cookie.path,
//         secure: cookie.secure,
//         httpOnly: cookie.httpOnly
//       };
      
//       if (cookie.expiry) {
//         var expiryDate = new Date(cookie.expiry * 1000);
//         result.sessionInfo.expiryDate = expiryDate;
//         result.sessionInfo.isExpired = expiryDate < new Date();
        
//         if (!result.sessionInfo.isExpired) {
//           var daysLeft = Math.floor((expiryDate - new Date()) / (1000 * 60 * 60 * 24));
//           result.sessionInfo.daysLeft = daysLeft;
//         }
//       }
//     }
    
//     // 有効期限チェック
//     if (cookie.expiry) {
//       if (cookie.expiry < now) {
//         result.expiredCookies.push({
//           name: cookie.name,
//           expiredSince: new Date(cookie.expiry * 1000)
//         });
//       } else {
//         result.validCookies.push(cookie.name);
//       }
//     } else {
//       // セッションクッキー（有効期限なし）も有効とみなす
//       result.validCookies.push(cookie.name);
//     }
//   }
  
//   // 結果の判定
//   if (result.expiredCookies.length > 0) {
//     result.message = "⚠️ 期限切れのクッキーがあります";
//     result.detailedMessage.push("期限切れ: " + result.expiredCookies.map(function(c) { return c.name; }).join(", "));
    
//     // 全てのクッキーが期限切れか確認
//     if (result.expiredCookies.length === result.totalCookies) {
//       result.message = "❌ 全てのクッキーが期限切れです";
//       result.detailedMessage.push("対処法: aucfan_sheet_writer.pyで再ログインが必要です");
//     }
//   } else if (!result.hasSessionCookie) {
//     result.message = "⚠️ セッションクッキーが見つかりません";
//     result.detailedMessage.push("認証に必要なセッションクッキーが不足しています");
//     result.detailedMessage.push("対処法: aucfan_sheet_writer.pyで再ログインしてください");
//   } else if (result.sessionInfo && result.sessionInfo.isExpired) {
//     result.message = "❌ セッションクッキーが期限切れです";
//     result.detailedMessage.push("セッション名: " + result.sessionInfo.name);
//     result.detailedMessage.push("期限切れ日時: " + result.sessionInfo.expiryDate.toLocaleString('ja-JP'));
//     result.detailedMessage.push("対処法: aucfan_sheet_writer.pyで再ログインしてください");
//   } else {
//     result.isValid = true;
//     result.message = "✅ クッキーは有効です";
//     result.detailedMessage.push("有効なクッキー数: " + result.validCookies.length + "/" + result.totalCookies);
    
//     if (result.sessionInfo && result.sessionInfo.daysLeft !== undefined) {
//       result.detailedMessage.push("セッション有効期限: あと" + result.sessionInfo.daysLeft + "日");
//     }
//   }
  
//   return result;
// }

// /**
//  * オークファンで商品を検索
//  * @param {string} keyword - 検索キーワード
//  * @param {Array} cookies - クッキー配列
//  * @returns {Array} 検索結果の配列
//  */
// function searchAucfanProducts(keyword, cookies) {
//   // クッキーチェック
//   if (!cookies || cookies.length === 0) {
//     console.error("❌ searchAucfanProducts: クッキーが提供されていません");
//     console.error("対処法: loadCookiesFromSheet()でクッキーを読み込んでから実行してください");
//     throw new Error("クッキーが必要です。先にログインしてください。");
//   }
  
//   var cookieHeader = formatCookieHeader(cookies);
//   var encodedKeyword = encodeURIComponent(keyword);
//   var searchUrl = 'https://pro.aucfan.com/search?q=' + encodedKeyword;
  
//   console.log('検索URL: ' + searchUrl);
  
//   var options = {
//     'method': 'get',
//     'headers': {
//       'Cookie': cookieHeader,
//       'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
//       'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
//       'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
//       'Referer': 'https://ssl.aucfan.com/aucfanpro/'
//     },
//     'muteHttpExceptions': true,
//     'followRedirects': false,
//     'timeout': 30 // 30秒のタイムアウト設定
//   };
  
//   try {
//     var response = UrlFetchApp.fetch(searchUrl, options);
//     var responseCode = response.getResponseCode();
    
//     console.log('レスポンスコード: ' + responseCode);
    
//     if (responseCode === 200) {
//       var html = response.getContentText();
      
//       // ログイン画面にリダイレクトされていないか確認
//       if (html.includes('login') && html.includes('password')) {
//         console.error('ログインが必要です');
//         return [];
//       }
      
//       // HTMLから商品情報を抽出
//       return parseAucfanSearchResults(html);
      
//     } else if (responseCode === 302 || responseCode === 301) {
//       console.error('リダイレクトされました。ログインが必要な可能性があります');
//       return [];
//     } else {
//       console.error('エラー: HTTPステータス ' + responseCode);
//       return [];
//     }
    
//   } catch (e) {
//     console.error('リクエストエラー: ' + e.toString());
//     return [];
//   }
// }

// /**
//  * オークファンの検索結果HTMLを解析
//  * @param {string} html - HTMLコンテンツ
//  * @returns {Array} 商品情報の配列
//  */
// function parseAucfanSearchResults(html) {
//   var results = [];
  
//   try {
//     // シンプルな正規表現でパース（実際のHTML構造に合わせて調整が必要）
//     // 商品リストのコンテナを探す
//     var itemPattern = /<div[^>]*class="[^"]*item[^"]*"[^>]*>[\s\S]*?<\/div>/gi;
//     var items = html.match(itemPattern) || [];
    
//     console.log('見つかった商品候補: ' + items.length);
    
//     // 各商品から情報を抽出
//     for (var i = 0; i < Math.min(items.length, 20); i++) { // 最大20件
//       var item = items[i];
//       var product = {};
      
//       // タイトルを抽出
//       var titleMatch = item.match(/<a[^>]*>([^<]+)<\/a>/);
//       if (titleMatch) {
//         product.title = titleMatch[1].trim();
//       }
      
//       // 価格を抽出
//       var priceMatch = item.match(/([0-9,]+)\s*円/);
//       if (priceMatch) {
//         product.price = priceMatch[1];
//       }
      
//       // 日付を抽出
//       var dateMatch = item.match(/(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2})/);
//       if (dateMatch) {
//         product.date = dateMatch[1];
//       }
      
//       // URLを抽出
//       var urlMatch = item.match(/<a[^>]*href="([^"]+)"/);
//       if (urlMatch) {
//         product.url = urlMatch[1];
//         if (!product.url.startsWith('http')) {
//           product.url = 'https://ssl.aucfan.com' + product.url;
//         }
//       }
      
//       // 有効なデータのみ追加
//       if (product.title || product.price) {
//         product.timestamp = new Date();
//         results.push(product);
//       }
//     }
    
//     // 代替パターン（構造が異なる場合）
//     if (results.length === 0) {
//       console.log('代替パターンで解析を試みます...');
      
//       // 価格パターンを探す
//       var pricePattern = /<span[^>]*class="[^"]*price[^"]*"[^>]*>([^<]+)<\/span>/gi;
//       var prices = html.match(pricePattern) || [];
      
//       // タイトルパターンを探す
//       var titlePattern = /<h\d[^>]*class="[^"]*title[^"]*"[^>]*>([^<]+)<\/h\d>/gi;
//       var titles = html.match(titlePattern) || [];
      
//       var count = Math.min(prices.length, titles.length, 20);
//       for (var j = 0; j < count; j++) {
//         results.push({
//           title: titles[j] ? titles[j].replace(/<[^>]+>/g, '').trim() : 'タイトル不明',
//           price: prices[j] ? prices[j].replace(/<[^>]+>/g, '').trim() : '価格不明',
//           timestamp: new Date()
//         });
//       }
//     }
    
//   } catch (e) {
//     console.error('HTML解析エラー: ' + e.toString());
//   }
  
//   console.log('解析結果: ' + results.length + '件');
//   return results;
// }

// /**
//  * オークファンの検索結果をシートに保存
//  * @param {Array} results - 検索結果
//  * @param {string} keyword - 検索キーワード
//  */
// function saveAucfanResultsToSheet(results, keyword) {
//   var sheetName = 'AucfanResults';
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
//   // シートがなければ作成
//   if (!sheet) {
//     sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    
//     // ヘッダーを設定
//     var headers = ['取得日時', '検索キーワード', 'タイトル', '価格', '日付', 'URL'];
//     sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
//     // ヘッダーのスタイルを設定
//     sheet.getRange(1, 1, 1, headers.length)
//       .setBackground('#4285F4')
//       .setFontColor('#FFFFFF')
//       .setFontWeight('bold');
//   }
  
//   // データを準備
//   var data = [];
//   for (var i = 0; i < results.length; i++) {
//     var result = results[i];
//     data.push([
//       result.timestamp || new Date(),
//       keyword,
//       result.title || '',
//       result.price || '',
//       result.date || '',
//       result.url || ''
//     ]);
//   }
  
//   // データを追加
//   if (data.length > 0) {
//     var lastRow = sheet.getLastRow();
//     sheet.getRange(lastRow + 1, 1, data.length, 6).setValues(data);
//     console.log(data.length + '件の結果を保存しました');
//   }
// }

// /**
//  * メニューに追加（onOpenに追加する）
//  */
// function addAucfanMenu() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('オークファン連携')
//     .addItem('固定URLでテスト実行', 'testAucfanPriceExtraction')
//     .addItem('テストスクレイピング実行', 'testAucfanScraping')
//     .addItem('ログイン状態確認', 'checkAucfanLoginStatus')
//     .addItem('クッキー診断', 'diagnoseCookies')
//     .addSeparator()
//     .addItem('結果シートを開く', 'openAucfanResultsSheet')
//     .addToUi();
// }

// /**
//  * オークファンのログイン状態を確認
//  */
// function checkAucfanLoginStatus() {
//   console.log("=== ログイン状態確認開始 ===");
//   var cookies = loadCookiesFromSheet();
  
//   if (cookies.length === 0) {
//     console.error('クッキーが見つかりません');
//     SpreadsheetApp.getActiveSpreadsheet().toast('クッキーが見つかりません', '❌ エラー', 5);
//     return;
//   }
  
//   console.log("クッキー数:", cookies.length);
  
//   // クッキーの有効期限チェック
//   var sessionCookie = cookies.find(function(c) { 
//     return c.name === 'PHPSESSID' || c.name.toLowerCase().includes('session'); 
//   });
  
//   if (sessionCookie && sessionCookie.expires) {
//     var expiryDate = new Date(sessionCookie.expires * 1000);
//     console.log("セッションクッキー有効期限:", expiryDate.toLocaleString());
    
//     if (expiryDate < new Date()) {
//       console.error("セッションクッキーの有効期限が切れています");
//       SpreadsheetApp.getActiveSpreadsheet().toast(
//         'セッションの有効期限切れです。再ログインが必要です。', 
//         '❌ セッション切れ', 
//         10
//       );
//       return;
//     }
//   }
  
//   var cookieHeader = formatCookieHeader(cookies);
//   var checkUrl = 'https://pro.aucfan.com/';
  
//   var options = {
//     'method': 'get',
//     'headers': {
//       'Cookie': cookieHeader,
//       'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
//       'Accept': 'text/html,application/xhtml+xml'
//     },
//     'muteHttpExceptions': true,
//     'followRedirects': false,
//     'timeout': 20 // 20秒のタイムアウト設定（ログイン確認なので短めに）
//   };
  
//   try {
//     var response = UrlFetchApp.fetch(checkUrl, options);
//     var responseCode = response.getResponseCode();
//     console.log("レスポンスコード:", responseCode);
    
//     if (responseCode === 200) {
//       var content = response.getContentText();
//       // ログインページかダッシュボードかを判定
//       if (content.includes('ログイン') && content.includes('パスワード')) {
//         console.error('ログインページが表示されました');
//         SpreadsheetApp.getActiveSpreadsheet().toast(
//           'セッション無効です。再ログインが必要です。', 
//           '❌ 要再ログイン', 
//           10
//         );
//       } else if (content.includes('マイページ') || content.includes('ログアウト')) {
//         console.log('ログイン状態: 有効');
//         SpreadsheetApp.getActiveSpreadsheet().toast(
//           'ログイン状態は有効です', 
//           '✅ ログイン中', 
//           5
//         );
//       } else {
//         console.log('ページ内容を確認できませんでした');
//         console.log('HTMLの一部:', content.substring(0, 300));
//         SpreadsheetApp.getActiveSpreadsheet().toast(
//           'ログイン状態を確認できませんでした', 
//           '⚠️ 不明', 
//           5
//         );
//       }
//     } else if (responseCode === 302 || responseCode === 301) {
//       var location = response.getHeaders()['Location'] || '';
//       console.error('リダイレクトされました:', location);
//       if (location.includes('login')) {
//         SpreadsheetApp.getActiveSpreadsheet().toast(
//           'ログインページにリダイレクトされました', 
//           '❌ 要ログイン', 
//           10
//         );
//       } else {
//         SpreadsheetApp.getActiveSpreadsheet().toast(
//           'リダイレクトされました: ' + location, 
//           '⚠️ リダイレクト', 
//           5
//         );
//       }
//     } else {
//       console.error('HTTPステータス:', responseCode);
//       SpreadsheetApp.getActiveSpreadsheet().toast(
//         'HTTPエラー: ' + responseCode, 
//         '❌ エラー', 
//         5
//       );
//     }
    
//   } catch (e) {
//     console.error('エラー:', e.toString());
//     SpreadsheetApp.getActiveSpreadsheet().toast(
//       'エラー: ' + e.toString(), 
//       '❌ エラー', 
//       10
//     );
//   }
  
//   console.log("=== ログイン状態確認終了 ===");
// }

// /**
//  * 結果シートを開く
//  */
// function openAucfanResultsSheet() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AucfanResults');
//   if (sheet) {
//     SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
//   } else {
//     SpreadsheetApp.getUi().alert('結果シートがまだ作成されていません');
//   }
// }

// /**
//  * スプレッドシートを開いたときに実行される関数
//  * オークファンメニューを追加
//  */
// function onOpen() {
//   addAucfanMenu();
// }
