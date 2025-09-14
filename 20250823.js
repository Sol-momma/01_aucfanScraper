// /** ================== 定数 ================== **/
// const COL_B = 2;              // B列
// const COL_C = 3;              // C列（出力起点）
// const ROW_HTML_START = 28;    // B28からHTML貼付開始
// const ROW_LABELS_START = 10;  // ラベル群の開始行（目安）
// const ROW_CLEAR_10 = 10;      // 10行目をクリア対象
// const ROW_CLEAR_11 = 11;      // 11行目をクリア対象
// const ROW_CLEAR_BLOCK_FROM = 14; // 14行目から
// const ROW_CLEAR_BLOCK_TO   = 20; // 20行目までをクリア対象（ショップを含む）
// const ROW_CLEAR_FROM       = 27; // 27行目以降クリア対象
// const MAX_OUTPUT_ITEMS = 20;  // C〜V（20列）に合わせる場合
// const LABELS = ['詳細URL','画像URL','日付','ランク','価格','ショップ','引用サイト']; // B10〜B17で使う想定のラベル
// const DO_AUTO_RESIZE = false;              // ← ここを false にしておく
// const APPLY_CLIP_WRAP = true;              // ← 文字折返しで高さ変化を防ぐなら true
// const MAX_OUTPUT_COLS = 20;                // C～V


// /** ============== 共通ユーティリティ ============== **/

// // ROW_HTML_START 以降に貼られたHTMLを結合して取得（改行で連結）
// // 楽天URLの場合は自動的にフェッチ
// function readHtmlFromB28_(sheet) {
//   var last = sheet.getLastRow();
//   if (last < ROW_HTML_START) throw new Error('B' + ROW_HTML_START + 'にHTMLがありません');
//   var vals = sheet.getRange(ROW_HTML_START, COL_B, last - ROW_HTML_START + 1, 1)
//                   .getValues()
//                   .map(function(r){ return r[0]; });
//   var joined = vals.filter(function(v){ return v !== '' && v !== null; }).join('\n');
//   if (!joined) throw new Error('B' + ROW_HTML_START + 'から下にHTMLが見つかりません');
  
//   // URLかどうかチェック
//   var trimmed = joined.trim();
//   if (trimmed.match(/^https?:\/\/.*rakuten\.co\.jp\//i)) {
//     return fetchRakutenHtml_(trimmed);
//   }
//   if (trimmed.match(/^https?:\/\/.*starbuyers-global-auction\.com\//i)) {
//     return fetchStarBuyersHtml_(trimmed);
//   }
  
//   return String(joined);
// }

// // 全角数字→半角、カンマ/空白除去
// function normalizeNumber_(s) {
//   if (s == null) return '';
//   s = String(s);
//   s = s.replace(/[！-～]/g, function(ch){ return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); });
//   s = s.replace(/\u3000/g, ' ');
//   s = s.replace(/[,\s]/g, '');
//   // 数字のみで構成されているかチェック
//   if (s && /^\d+$/.test(s)) {
//     return s;
//   }
//   return ''; // 数値でない場合は空文字列を返す
// }

// // 楽天URLからHTMLをフェッチ
// function fetchRakutenHtml_(url) {
//   try {
//     var options = {
//       method: 'get',
//       headers: {
//         'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
//       },
//       muteHttpExceptions: true
//     };
    
//     var response = UrlFetchApp.fetch(url, options);
//     var html = response.getContentText('UTF-8');
    
//     if (response.getResponseCode() !== 200) {
//       throw new Error('楽天からのHTMLの取得に失敗しました。ステータスコード: ' + response.getResponseCode());
//     }
    
//     return html;
//   } catch (e) {
//     throw new Error('楽天URLからのHTMLフェッチエラー: ' + e.message);
//   }
// }

// // Star Buyers AuctionのURLからHTMLをフェッチ（ログイン必要）
// function fetchStarBuyersHtml_(url) {
//   try {
//     // ランダムな待機時間（1-3秒）
//     Utilities.sleep(1000 + Math.floor(Math.random() * 2000));
    
//     // ログイン情報（ベタ打ち）
//     var loginEmail = 'inui.hur@gmail.com';
//     var loginPassword = 'hur22721';
    
//     // ログインURL
//     var loginUrl = 'https://www.starbuyers-global-auction.com/login';
    
//     // ステップ1: ログインページにアクセスしてCSRFトークンとクッキーを取得
//     var loginPageOptions = {
//       method: 'get',
//       headers: {
//         'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
//         'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
//         'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
//         'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
//         'sec-ch-ua-mobile': '?0',
//         'sec-ch-ua-platform': '"macOS"',
//         'sec-fetch-dest': 'document',
//         'sec-fetch-mode': 'navigate',
//         'sec-fetch-site': 'none',
//         'sec-fetch-user': '?1',
//         'upgrade-insecure-requests': '1'
//       },
//       muteHttpExceptions: true,
//       followRedirects: false
//     };
    
//     var loginPageResponse = UrlFetchApp.fetch(loginUrl, loginPageOptions);
//     var loginPageHtml = loginPageResponse.getContentText('UTF-8');
//     var cookies = loginPageResponse.getAllHeaders()['Set-Cookie'];
    
//     // CSRFトークンを抽出（フォーム内の_tokenまたはmeta tagから）
//     var csrfToken = '';
//     var csrfMatch = loginPageHtml.match(/<input[^>]+name="_token"[^>]+value="([^"]+)"/);
//     if (!csrfMatch) {
//       csrfMatch = loginPageHtml.match(/<meta name="csrf-token" content="([^"]+)"/);
//     }
//     if (csrfMatch) {
//       csrfToken = csrfMatch[1];
//     }
    
//     console.log('CSRFトークン取得:', csrfToken ? '成功' : '失敗');
    
//     // ログインフォームのアクションURLを確認
//     var formActionMatch = loginPageHtml.match(/<form[^>]+action="([^"]+)"[^>]*>/);
//     var actualLoginUrl = loginUrl;
//     if (formActionMatch) {
//       var action = formActionMatch[1];
//       if (action.startsWith('/')) {
//         actualLoginUrl = 'https://www.starbuyers-global-auction.com' + action;
//       } else if (action.startsWith('http')) {
//         actualLoginUrl = action;
//       }
//     }
//     console.log('実際のログインURL:', actualLoginUrl);
    
//     // ログインフォームのフィールドを確認
//     var emailFieldMatch = loginPageHtml.match(/<input[^>]+type="email"[^>]+name="([^"]+)"/);
//     var emailFieldName = emailFieldMatch ? emailFieldMatch[1] : 'email';
//     console.log('Emailフィールド名:', emailFieldName);
    
//     // セッションクッキーを収集
//     var cookieMap = {};
//     if (cookies) {
//       if (Array.isArray(cookies)) {
//         cookies.forEach(function(cookie) {
//           var parts = cookie.split(';')[0].split('=');
//           if (parts.length >= 2) {
//             cookieMap[parts[0]] = parts.slice(1).join('=');
//           }
//         });
//       } else {
//         var parts = cookies.split(';')[0].split('=');
//         if (parts.length >= 2) {
//           cookieMap[parts[0]] = parts.slice(1).join('=');
//         }
//       }
//     }
    
//     console.log('取得したクッキー:', Object.keys(cookieMap).join(', '));
    
//     // ランダムな待機時間（1-2秒）
//     Utilities.sleep(1000 + Math.floor(Math.random() * 1000));
    
//     // ステップ2: ログイン実行
//     var loginPayload = {};
//     loginPayload[emailFieldName] = loginEmail;
//     loginPayload['password'] = loginPassword;
//     loginPayload['_token'] = csrfToken;
    
//     // Remember meチェックボックスがある場合
//     if (loginPageHtml.indexOf('name="remember"') > -1) {
//       loginPayload['remember'] = '1';
//     }
    
//     // クッキー文字列を構築
//     var cookieString = Object.keys(cookieMap).map(function(key) {
//       return key + '=' + cookieMap[key];
//     }).join('; ');
    
//     // XSRF-TOKENがある場合、X-XSRF-TOKENヘッダーとして追加
//     var xsrfToken = cookieMap['XSRF-TOKEN'] || '';
//     if (xsrfToken) {
//       // URLデコード（LaravelのXSRF-TOKENは通常URLエンコードされている）
//       try {
//         xsrfToken = decodeURIComponent(xsrfToken);
//       } catch(e) {}
//     }
    
//     var loginHeaders = {
//       'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
//       'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
//       'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
//       'Content-Type': 'application/x-www-form-urlencoded',
//       'Cookie': cookieString,
//       'Referer': loginUrl,
//       'Origin': 'https://www.starbuyers-global-auction.com',
//       'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
//       'sec-ch-ua-mobile': '?0',
//       'sec-ch-ua-platform': '"macOS"',
//       'sec-fetch-dest': 'document',
//       'sec-fetch-mode': 'navigate',
//       'sec-fetch-site': 'same-origin',
//       'sec-fetch-user': '?1',
//       'upgrade-insecure-requests': '1'
//     };
    
//     if (xsrfToken) {
//       loginHeaders['X-XSRF-TOKEN'] = xsrfToken;
//     }
    
//     var payloadString = Object.keys(loginPayload).map(function(key) {
//       return encodeURIComponent(key) + '=' + encodeURIComponent(loginPayload[key]);
//     }).join('&');
    
//     console.log('ログインペイロード:', payloadString.replace(loginPassword, '***'));
    
//     var loginOptions = {
//       method: 'post',
//       headers: loginHeaders,
//       payload: payloadString,
//       muteHttpExceptions: true,
//       followRedirects: false
//     };
    
//     var loginResponse = UrlFetchApp.fetch(actualLoginUrl, loginOptions);
//     console.log('ログインレスポンスステータス:', loginResponse.getResponseCode());
    
//     // ログイン後のクッキーを更新
//     var loginCookies = loginResponse.getAllHeaders()['Set-Cookie'];
//     if (loginCookies) {
//       if (Array.isArray(loginCookies)) {
//         loginCookies.forEach(function(cookie) {
//           var parts = cookie.split(';')[0].split('=');
//           if (parts.length >= 2) {
//             cookieMap[parts[0]] = parts.slice(1).join('=');
//           }
//         });
//       } else {
//         var parts = loginCookies.split(';')[0].split('=');
//         if (parts.length >= 2) {
//           cookieMap[parts[0]] = parts.slice(1).join('=');
//         }
//       }
//     }
    
//     // リダイレクト先を確認
//     var locationHeader = loginResponse.getAllHeaders()['Location'] || loginResponse.getAllHeaders()['location'];
//     console.log('リダイレクト先:', locationHeader || 'なし');
//     if (locationHeader) {
//       // リダイレクト先にアクセス
//       cookieString = Object.keys(cookieMap).map(function(key) {
//         return key + '=' + cookieMap[key];
//       }).join('; ');
      
//       var redirectOptions = {
//         method: 'get',
//         headers: {
//           'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
//           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
//           'Cookie': cookieString
//         },
//         muteHttpExceptions: true
//       };
      
//       var redirectResponse = UrlFetchApp.fetch(locationHeader, redirectOptions);
//       var redirectCookies = redirectResponse.getAllHeaders()['Set-Cookie'];
//       if (redirectCookies) {
//         if (Array.isArray(redirectCookies)) {
//           redirectCookies.forEach(function(cookie) {
//             var parts = cookie.split(';')[0].split('=');
//             if (parts.length >= 2) {
//               cookieMap[parts[0]] = parts.slice(1).join('=');
//             }
//           });
//         }
//       }
//     }
    
//     // ランダムな待機時間（2-4秒）
//     Utilities.sleep(2000 + Math.floor(Math.random() * 2000));
    
//     // ステップ3: 目的のURLにアクセス
//     cookieString = Object.keys(cookieMap).map(function(key) {
//       return key + '=' + cookieMap[key];
//     }).join('; ');
    
//     var fetchOptions = {
//       method: 'get',
//       headers: {
//         'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
//         'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
//         'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
//         'Cookie': cookieString,
//         'Referer': 'https://www.starbuyers-global-auction.com/',
//         'sec-ch-ua': '"Not;A=Brand";v="99", "Google Chrome";v="139", "Chromium";v="139"',
//         'sec-ch-ua-mobile': '?0',
//         'sec-ch-ua-platform': '"macOS"',
//         'sec-fetch-dest': 'document',
//         'sec-fetch-mode': 'navigate',
//         'sec-fetch-site': 'none',
//         'sec-fetch-user': '?1',
//         'upgrade-insecure-requests': '1'
//       },
//       muteHttpExceptions: true
//     };
    
//     var response = UrlFetchApp.fetch(url, fetchOptions);
//     var html = response.getContentText('UTF-8');
    
//     if (response.getResponseCode() !== 200) {
//       throw new Error('Star Buyersからのデータ取得に失敗しました。ステータスコード: ' + response.getResponseCode());
//     }
    
//     // ログインが必要なページにリダイレクトされていないかチェック
//     if (html.indexOf('Login') > -1 && html.indexOf('E-Mail Address') > -1 && html.indexOf('p-item-list__body') === -1) {
//       console.log('HTMLレスポンスの一部:', html.substring(0, 500));
//       throw new Error('ログインに失敗しました。認証情報を確認してください。');
//     }
    
//     // 商品データが含まれているかチェック
//     if (html.indexOf('p-item-list__body') === -1) {
//       console.log('取得したHTMLに商品データが含まれていない可能性があります。');
//     }
    
//     return html;
//   } catch (e) {
//     throw new Error('Star Buyers URLからのHTMLフェッチエラー: ' + e.message);
//   }
// }

// // HTMLエンティティ最小限デコード
// function htmlDecode_(s) {
//   if (!s) return s;
//   return s
//     .replace(/&amp;/g, '&')
//     .replace(/&lt;/g, '<')
//     .replace(/&gt;/g, '>')
//     .replace(/&quot;/g, '"')
//     .replace(/&#039;/g, "'")
//     .replace(/&nbsp;/g, ' ');
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
//   return bVals.map(function(r){ return String(r[0] || '').trim(); });
// }

// // ラベル→横並びデータを書き込み（C列起点）
// function writeRowByLabel_(sheet, bHeadings, label, rowValuesArray) {
//   var idx = bHeadings.findIndex(function(v){ return v === label; });
//   if (idx === -1) return; // 見出し未設置ならスキップ
//   var row = idx + 1;
//   clearRowFromColumnC_(sheet, row);
//   if (!rowValuesArray || !rowValuesArray.length) return;

//   // 件数制限（C〜Vの20件以内）
//   var out = rowValuesArray.slice(0, MAX_OUTPUT_ITEMS);
//   var values2d = [ out ];
//   sheet.getRange(row, COL_C, 1, out.length).setValues(values2d);
// }

// /** ============== ソース判定 ============== **/
// function detectSource_(html) {
//   var h = html.toLowerCase();
//   if (h.indexOf('starbuyers-global-auction.com') !== -1 ||
//       h.indexOf('star buyers auction') !== -1 ||
//       h.indexOf('p-item-list__body__cell') !== -1) {
//     return 'starbuyers';
//   }
//   if (h.indexOf('aucfan') !== -1 ||
//       h.indexOf('オークファン') !== -1 ||
//       h.indexOf('落札相場') !== -1) {
//     return 'aucfan';
//   }
//   if (h.indexOf('rakuten.co.jp') !== -1 ||
//       h.indexOf('楽天市場') !== -1 ||
//       h.indexOf('送料無料') !== -1 ||
//       h.indexOf('ポイント(1倍)') !== -1) {
//     return 'rakuten';
//   }
//   return 'starbuyers';
// }
// // --- 置き換え版：STAR BUYERS パーサ ---
// // 返却：[{detailUrl, imageUrl, date, rank, price}]
// function parseStarBuyersFromHtml_(html) {
//   var items = [];
//   var H = String(html);

//   // 1) 各アイテムの開始インデックスを列挙
//   var starts = [];
//   var re = /<div\s+class="p-item-list__body">/g;
//   var m;
//   while ((m = re.exec(H)) !== null) {
//     starts.push(m.index);
//   }
//   if (starts.length === 0) return items;

//   // 2) 開始インデックスごとに「次の開始 or 文末」までを1件ブロックとする
//   for (var i = 0; i < starts.length; i++) {
//     var s = starts[i];
//     var e = (i + 1 < starts.length) ? starts[i + 1] : H.length;
//     var block = H.slice(s, e);

//     // 詳細URL
//     var detailUrl = '';
//     var mUrl = block.match(/<a[^>]+href="(https:\/\/www\.starbuyers-global-auction\.com\/market_price\/[^"]+)"/i);
//     if (mUrl) detailUrl = htmlDecode_(mUrl[1]);

//     // 画像URL
//     var imageUrl = '';
//     var mImg = block.match(/<img[^>]+src="([^"]+)"[^>]*>/i);
//     if (mImg) imageUrl = htmlDecode_(mImg[1]);

//     // 開催日
//     var date = '';
//     var mDate = block.match(/data-head="開催日"[\s\S]*?<strong>([^<]+)<\/strong>/i);
//     if (mDate) date = mDate[1].trim();

//     // ランク（data-rank="ＡＢ" など）→ 許可値に正規化
//     var rank = '';
//     var mRank = block.match(/data-rank="([^"]+)"/i);
//     if (mRank) rank = normalizeRank_(mRank[1].trim());

//     // 価格（"96,000yen" 等）
//     var price = '';
//     var mPrice = block.match(/data-head="落札額"[\s\S]*?<strong>([^<]+)<\/strong>/i);
//     if (mPrice) {
//       var raw = mPrice[1].trim();
//       var num = normalizeNumber_(raw.replace(/yen/i,''));
//       // 数値化できた場合は数値、できなかった場合は元のテキストを保持
//       price = num || raw;
//       console.log('SBA価格抽出:', raw, '→', price);
//     }

//     if (detailUrl || imageUrl || date || rank || price) {
//       items.push({ 
//         detailUrl: detailUrl, 
//         imageUrl: imageUrl, 
//         date: date, 
//         rank: rank, 
//         price: price,
//         shop: '',
//         source: 'SBA'
//       });
//     }
//   }

//   return items;
// }

// /** オークファン用パーサ（元の extractAucfanItems を簡略化して利用） */
// function parseAucfanFromHtml_(html) {
//   const items = [];
//   const base = 'https://pro.aucfan.com';
//   const re = /<li class="item">\s*<ul>([\s\S]*?)<\/ul>\s*<\/li>/gi;
//   let m;
//   while ((m = re.exec(html)) !== null) {
//     const block = m[1];
//     const detailPath = firstMatch_(block, /<li class="title">[\s\S]*?<a\s+href="([^"]+)"/i);
//     const detailUrl = detailPath ? (detailPath.startsWith('http') ? detailPath : base + detailPath) : '';
//     const imageUrl = firstMatch_(block, /<img[^>]+class="thumbnail"[^>]+src="([^"]+)"/i);
//     const price = normalizeNumber_(firstMatch_(block, /<li class="price"[^>]*>\s*([^<]+)/i));
//     const endTxt = firstMatch_(block, /<li class="end">\s*([^<]+)/i);
//     // endTxtそのまま日付に入れる（必要なら deriveEndDateValue_ を再利用可）

//     if (detailUrl || imageUrl || price || endTxt) {
//       items.push({
//         detailUrl: detailUrl,
//         imageUrl: imageUrl,
//         date: endTxt || '',
//         rank: '', // aucfanにはランクが基本出ないので空
//         price: price || '',
//         shop: '',
//         source: 'オークファン'
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
//     /<li[^>]*>[\s\S]*?<a[^>]+href="https?:\/\/item\.rakuten\.co\.jp[^>]*>[\s\S]*?<\/li>/gi
//   ];
  
//   let productBlocks = [];
  
//   // 各パターンを試して商品ブロックを抽出
//   for (let pattern of blockPatterns) {
//     const matches = H.match(pattern);
//     if (matches && matches.length > 0) {
//       productBlocks = matches;
//       console.log('楽天商品ブロック検出:', matches.length + '件（パターン:', blockPatterns.indexOf(pattern) + '）');
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
//       const detailUrlMatch = block.match(/<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/i);
//       if (!detailUrlMatch) continue;
      
//       const detailUrl = htmlDecode_(detailUrlMatch[1]);
//       const testItem = extractRakutenItemFromBlock_(block, detailUrl, i);
//       testItems.push(testItem);
      
//       // 価格が見つからない場合
//       if (!testItem.price) {
//         console.log('ブロック方式で価格が見つかりません。リンクベース方式に切り替えます。');
//         useBlockMethod = false;
//         break;
//       }
//     }
    
//     if (useBlockMethod) {
//       // ブロック方式で全商品を処理
//       productBlocks.forEach(function(block, index) {
//         const detailUrlMatch = block.match(/<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/i);
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
//     console.log('商品ブロックが見つかりません。リンクベースの抽出に切り替えます。');
//     const linkRe = /<a[^>]+href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"[^>]*>/gi;
//     let links = [];
//     let m;
//     while ((m = linkRe.exec(H)) !== null) {
//       links.push({
//         url: m[1],
//         startIndex: m.index,
//         endIndex: m.index + m[0].length
//       });
//     }
    
//     // 各リンクに対して、周辺の情報を抽出
//     links.forEach(function(link, index) {
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
//     console.log('=== 楽天商品ブロック ' + (index + 1) + ' ===');
//     console.log('URL:', detailUrl);
//     console.log('ブロック長:', block.length);
    
//     // 実際の商品価格を探すためのデバッグ
//     const allPriceMatches = block.match(/[0-9,]+\s*円/g);
//     if (allPriceMatches) {
//       console.log('ブロック内の全価格候補:', allPriceMatches.filter(p => !p.includes('3,980')).slice(0, 5));
//     }
    
//     // 価格タグを探す
//     const priceTagMatches = block.match(/<(?:span|div|p)[^>]*(?:class|id)="[^"]*price[^"]*"[^>]*>[^<]+<\/(?:span|div|p)>/gi);
//     if (priceTagMatches) {
//       console.log('価格タグ数:', priceTagMatches.length);
//       if (priceTagMatches[0]) {
//         console.log('最初の価格タグ:', priceTagMatches[0].replace(/\s+/g, ' '));
//       }
//     }
    
//     // 商品リンクの後の価格を探す
//     const afterLinkMatch = block.match(/item\.rakuten\.co\.jp[^>]+>[\s\S]{0,300}?([0-9,]+)\s*円/);
//     if (afterLinkMatch) {
//       console.log('商品リンク後の価格:', afterLinkMatch[1] + '円');
//     }
//   }
    
//     // 画像URL（リンクに最も近いimg要素）
//     let imageUrl = '';
//     const imgMatches = block.match(/<img[^>]+src="([^"]+)"[^>]*>/gi);
//     if (imgMatches && imgMatches.length > 0) {
//       // 最初に見つかった画像を使用
//       const imgMatch = imgMatches[0].match(/src="([^"]+)"/i);
//       if (imgMatch) {
//         imageUrl = htmlDecode_(imgMatch[1]);
//       }
//     }
    
//     // 価格（円を含む数値、カンマ付き）- 複数の価格から最小値を選択
//     let price = '';
//     let prices = [];
    
//       // 価格検出の前に、送料無料ラインの説明文を除去
//   let priceSearchBlock = block;
//   // 送料無料ラインの説明文を除去（3980を誤検出しないため）
//   priceSearchBlock = priceSearchBlock.replace(/送料無料ラインを[^。]+。/g, '');
//   priceSearchBlock = priceSearchBlock.replace(/3,980円以下に設定[^。]+。/g, '');
  
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
//     /item\.rakuten\.co\.jp[^>]*>[\s\S]{0,500}?([0-9,]+)\s*円/gi
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
//       if (normalized && normalized !== '0') {
//         const numPrice = parseInt(normalized);
//         if (!isNaN(numPrice) && numPrice >= 100) { // 100円以上を有効な価格とする
//           prices.push(numPrice);
//           if (index < 3) {
//             console.log('価格エリアから発見:', priceMatch[1], '→', numPrice);
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
//       /¥([0-9]{1,3}(?:,[0-9]{3})*)/g
//     ];
    
//     // 全ての価格パターンで価格を収集
//     for (let pattern of pricePatterns) {
//       let match;
//       const globalPattern = new RegExp(pattern.source, 'gi');
//       while ((match = globalPattern.exec(priceSearchBlock)) !== null) {
//         const rawPrice = match[1];
//         const normalized = normalizeNumber_(rawPrice);
//         if (normalized && normalized !== '0') {
//           const numPrice = parseInt(normalized);
//           if (!isNaN(numPrice) && numPrice >= 100) {
//             prices.push(numPrice);
//             if (index < 3) {
//               console.log('パターンから発見:', rawPrice, '→', numPrice);
//             }
//           }
//         }
//       }
//     }
//   }
    
//     // 送料込みフラグを確認（送料込みの価格は除外の候補）
//     const hasFreeShipping = block.indexOf('送料無料') > -1 || block.indexOf('送料込') > -1;
    
//     if (prices.length > 0) {
//       // 重複を除去してソート
//       prices = [...new Set(prices)].sort((a, b) => a - b);
      
//       if (index < 3) {
//         console.log('見つかった価格リスト:', prices);
//       }
      
//       // 価格選択ロジック（改善版）
//       if (prices.length === 1) {
//         // 価格が1つだけの場合はそれを使用
//         price = String(prices[0]);
//       } else if (prices.length > 1) {
//         // 複数の価格がある場合の選択ロジック
        
//         // Step 1: 3980円（送料無料ライン）を除外
//         let candidates = prices.filter(p => p !== 3980);
        
//         if (candidates.length === 0) {
//           // 全て3980の場合は仕方なく使用
//           candidates = prices;
//         }
        
//         // Step 2: 送料と思われる価格を識別して除外
//         // 一般的な送料: 500-1000円
//         const shippingPrices = [500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000];
//         const nonShippingPrices = candidates.filter(p => !shippingPrices.includes(p));
        
//         if (nonShippingPrices.length > 0) {
//           // 送料以外の価格がある場合
//           // 1000円以上で最も小さい価格を選択（送料込み価格を避ける）
//           const over1000 = nonShippingPrices.filter(p => p >= 1000);
//           if (over1000.length > 0) {
//             price = String(Math.min(...over1000));
//           } else {
//             // 1000円未満しかない場合は最大値を選択
//             price = String(Math.max(...nonShippingPrices));
//           }
//         } else {
//           // 全て送料と思われる価格の場合
//           // 1000円以上の価格があればその最小値を選択
//           const reasonablePrices = candidates.filter(p => p >= 1000);
//           if (reasonablePrices.length > 0) {
//             price = String(Math.min(...reasonablePrices));
//           } else {
//             // 全て1000円未満の場合は最大値を使用
//             price = String(Math.max(...candidates));
//           }
//         }
        
//         // デバッグ情報
//         if (index < 3) {
//           console.log('価格選択結果:', price, '（候補:', candidates.join(','), '）');
//         }
//       }
//     }
    
//     // 価格が見つからない場合のフォールバック
//     if (!price) {
//       // より広範囲で価格を探す（最初に見つかった数値+円）
//       const fallbackMatch = block.match(/([0-9,]+)\s*円/);
//       if (fallbackMatch) {
//         const rawPrice = fallbackMatch[1];
//         const normalized = normalizeNumber_(rawPrice);
//         // 数値化できた場合は数値、できなかった場合は元のテキストを保持
//         price = normalized || rawPrice;
//       }
//     }
    
//     // デバッグ：価格抽出結果
//     if (index < 3) {
//       console.log('抽出された価格:', price || '価格が見つかりません');
//     }
    
//     // ショップ名の抽出（複数のパターンで試行）
//     let shop = '';
    
//     // パターン1: 楽天の一般的なショップ名表示パターン（改善版）
//     const shopPatterns = [
//       // 楽天市場店を含むリンク（最も確実）
//       /<a[^>]*>([^<]*楽天市場店)<\/a>/,
//       // 【】を含むショップ名（銀座パリスなど）
//       /<a[^>]*>(【[^】]+】[^<]*)<\/a>/,
//       // 単純なリンクで「店」を含む
//       /<a[^>]*>([^<]*店[^<]*)<\/a>/,
//       // merchantクラスを含む要素
//       /<(?:a|span|div)[^>]+class="[^"]*merchant[^"]*"[^>]*>([^<]+)<\/(?:a|span|div)>/i,
//       // 送料の前のリンク（シンプル版）
//       /<a[^>]*>([^<]+)<\/a>[^<]*送料/,
//       // content-merchant内のテキスト
//       /<div[^>]*class="[^"]*content-merchant[^"]*"[^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>/i,
//       // shopやstoreを含むhref属性のリンク
//       /<a[^>]*href="[^"]*\/(?:shop|store)\/[^"]*"[^>]*>([^<]+)<\/a>/,
//       // 価格の後のリンク（広範囲）
//       /円[^<]*<\/[^>]+>[^<]*<a[^>]*>([^<]+)<\/a>/,
//       // dui-linkbox内のショップリンク  
//       /<(?:a|span)[^>]*class="[^"]*dui-linkbox[^"]*"[^>]*>([^<]+)<\/(?:a|span)>/i,
//       // 39ショップの前のリンク
//       /<a[^>]*>([^<]+)<\/a>[^<]*39ショップ/,
//       // ポイントの前のリンク
//       /<a[^>]*>([^<]+)<\/a>[^<]*ポイント/,
//       // merchantやshop-nameクラス
//       /<[^>]+class="[^"]*(?:merchant-name|shop-name|store-name)[^"]*"[^>]*>([^<]+)<\/[^>]+>/i
//     ];
    
//     // 各パターンを試行
//     for (let i = 0; i < shopPatterns.length; i++) {
//       const pattern = shopPatterns[i];
//       const shopMatch = block.match(pattern);
//       if (shopMatch && shopMatch[1]) {
//         let candidate = shopMatch[1].trim();
        
//         // HTMLエンティティをデコード
//         candidate = htmlDecode_(candidate);
        
//         if (index < 3) {
//           console.log('パターン' + i + 'でマッチ:', candidate);
//         }
        
//         // ショップ名の妥当性チェック（緩和版）
//         if (candidate && 
//             candidate.length > 1 && // 1文字は除外
//             candidate.length < 100 && // ショップ名は100文字以内
//             candidate !== '">' && // ">のみは除外
//             !candidate.match(/^["><!-]+$/) && // HTMLタグの断片のみは除外
//             !candidate.match(/^[\d\s,]+$/) && // 数字、空白、カンマのみは除外
//             !candidate.startsWith('">') && // ">で始まるものは除外
//             !candidate.includes('<!--') && // HTMLコメントは除外
//             !candidate.match(/^(送料|ポイント|円|価格|個数|在庫|税込|税別|込み|無料|別|もっと見る|詳細を見る|レビュー|カート|かごに入れる)$/i)) { // 価格関連の単語のみは除外
          
//           // 日本語、英字、【】、店を含むものはOK
//           if (candidate.match(/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\u3400-\u4DBFa-zA-Z【】]/) || 
//               candidate.includes('店')) {
//             shop = candidate;
//             if (index < 3) {
//               console.log('採用されたショップ名:', candidate, '- パターン:', i);
//             }
//             break;
//           }
//         }
        
//         if (index < 3 && !shop && candidate) {
//           console.log('除外されたショップ名候補:', candidate, '- パターン:', i);
//         }
//       }
//     }
    
//     // パターン2: URLから店舗名を抽出（フォールバック）
//     if (!shop || shop.length === 0 || shop === '">') {
//       const shopFromUrl = detailUrl.match(/item\.rakuten\.co\.jp\/([^\/]+)\//);
//       if (shopFromUrl) {
//         shop = shopFromUrl[1];
//         if (index < 3) {
//           console.log('URLからショップ名を取得:', shop);
//         }
//       }
//     }
    
//     // 最終的な妥当性チェック
//     if (shop === '">' || shop.match(/^[">]+$/)) {
//       shop = '';
//     }
    
//     if (index < 3) {
//       console.log('最終的なショップ名:', shop || 'ショップ名が見つかりません');
//       if (!shop || shop === '">') {
//         // ショップ名が見つからない場合、リンクを含むエリアを表示
//         const linkArea = block.match(/<a[^>]+>([^<]{1,50})<\/a>[\s\S]{0,50}(?:送料|ポイント|39ショップ)/g);
//         if (linkArea) {
//           console.log('リンクエリア（最初の3つ）:');
//           linkArea.slice(0, 3).forEach((area, i) => {
//             console.log(i+1 + ':', area.replace(/\s+/g, ' ').substring(0, 100));
//           });
//         }
//       }
//     }
    
//     // 商品情報を返す
//     return {
//       detailUrl: detailUrl,
//       imageUrl: imageUrl,
//       date: '', // 楽天には日付情報がないため空
//       rank: '', // 楽天にはランク情報がないため空
//       price: price,
//       shop: shop,
//       source: '楽天'
//     };
// }

// /** 正規表現ユーティリティ */
// function firstMatch_(text, re) {
//   const m = re.exec(text);
//   return m ? (m[1] || '').trim() : '';
// }
// function writeItemsToTemplate_B10_V19_(sheet, items) {
//   var labels = ['詳細URL','画像URL','日付','ランク','価格','ショップ','引用サイト'];
//   var bHeadings = readColumnBHeadings_(sheet);

//   // まず通常整形
//   var rowArrays = {
//     '詳細URL': items.map(it => it.detailUrl || ''),
//     '画像URL': items.map(it => it.imageUrl || ''),
//     '日付'  : items.map(it => it.date || ''),
//     'ランク': items.map(it => it.rank || ''),
//     '価格'  : items.map(it => it.price || ''),
//     'ショップ': items.map(it => it.shop || ''),
//     '引用サイト': items.map(it => it.source || '')
//   };

//   // “ランク”は最終チェックして許可外は '' にする
//   if (rowArrays['ランク']) {
//     rowArrays['ランク'] = rowArrays['ランク'].map(normalizeRank_);
//   }

//   labels.forEach(function(label){
//     var rowIndex = bHeadings.indexOf(label);
//     if (rowIndex !== -1) {
//       console.log('ラベル「' + label + '」を行' + (rowIndex + 1) + 'に出力');
//       console.log('データ数:', rowArrays[label].length);
//       if (label === '価格' && rowArrays[label].length > 0) {
//         console.log('価格データの最初の5件:', rowArrays[label].slice(0, 5));
//       }
//       writeRowByLabel_(sheet, bHeadings, label, rowArrays[label]);
//       // （前回の対策）折返し抑止などはそのまま
//       if (APPLY_CLIP_WRAP) {
//         var row = rowIndex + 1;
//         var outLen = Math.min((items || []).length, MAX_OUTPUT_COLS);
//         if (outLen > 0) {
//           sheet.getRange(row, 3, 1, outLen)
//                .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
//         }
//       }
//     } else {
//       console.log('ラベル「' + label + '」がB列に見つかりません');
//     }
//   });

//   if (DO_AUTO_RESIZE && items.length > 0) {
//     sheet.autoResizeColumns(3, Math.max(1, Math.min(items.length, MAX_OUTPUT_COLS)));
//   }
// }

// // 全角→半角（英数のみ）
// function toHalfWidthAlphaNum_(s) {
//   return String(s || '').replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
// }

// // 許可ランクへ正規化：S, SA, A, AB, B, BC, C 以外は '' にする
// function normalizeRank_(rank) {
//   let r = toHalfWidthAlphaNum_(rank).toUpperCase().replace(/\s+/g,'');
//   // よくある表記ゆれ
//   const map = { 'Ｓ': 'S', 'ＳＡ': 'SA', 'Ａ': 'A', 'ＡＢ': 'AB', 'Ｂ': 'B', 'ＢＣ': 'BC', 'Ｃ': 'C' };
//   if (map[rank]) r = map[rank]; // 全角キーへの保険
//   // 2文字以上の全角が混在していたケースも救済
//   r = r.replace('Ａ','A').replace('Ｂ','B').replace('Ｃ','C').replace('Ｓ','S');
//   const ALLOWED = new Set(['S','SA','A','AB','B','BC','C']);
//   return ALLOWED.has(r) ? r : '';
// }

// // D2の査定ジャンルに基づいてB16（服サイズ）を自動設定
// // D2に「服」が含まれていれば空白、含まれていなければ「なし」を設定
// function handleClothingSizeField_(sheet) {
//   // D2セルの値を取得（査定ジャンル）
//   var categoryValue = sheet.getRange('D2').getValue();
//   if (!categoryValue) return;
  
//   var categoryStr = String(categoryValue);
  
//   // B列の見出しを取得
//   var bHeadings = readColumnBHeadings_(sheet);
  
//   // 「服サイズ」の行を探す（通常はB16）
//   var clothingSizeRow = -1;
//   for (var i = 0; i < bHeadings.length; i++) {
//     if (bHeadings[i] === '服サイズ') {
//       clothingSizeRow = i + 1;
//       break;
//     }
//   }
  
//   if (clothingSizeRow === -1) return; // 服サイズ行が見つからない
  
//   // C〜V列（20列）に値を設定
//   var valueToSet = '';
//   if (categoryStr.indexOf('服') !== -1) {
//     // 「服」が含まれている場合は空白（何も入力しない）
//     valueToSet = '';
//     console.log('D2に「服」が含まれているため、服サイズを空白に設定');
//   } else {
//     // 「服」が含まれていない場合は「なし」を設定
//     valueToSet = 'なし';
//     console.log('D2に「服」が含まれていないため、服サイズを「なし」に設定');
//   }
  
//   // C列からV列まで（20列）同じ値を設定
//   var values = [];
//   for (var i = 0; i < MAX_OUTPUT_ITEMS; i++) {
//     values.push(valueToSet);
//   }
//   sheet.getRange(clothingSizeRow, COL_C, 1, MAX_OUTPUT_ITEMS).setValues([values]);
// }

// /** ============== URL自動生成機能 ============== **/
// // J6、J7、J8を読み取って、B28に検索URLを自動生成
// // J6: サイト種別（「楽天」または「SBA」）
// // J7: 検索キーワード（例：「プラダ テスート サフィアーノ」）
// // J8: 期間（SBAのみ、日付またはYYYY-MM-DD形式の文字列、例：「2025-02-01」）
// function generateSearchUrl() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getActiveSheet();
  
//   // J6: サイト種別（楽天/SBA）、J7: フリーワード、J8: 期間
//   var siteType = sheet.getRange('J6').getValue();
//   var keyword = sheet.getRange('J7').getValue();
//   var period = sheet.getRange('J8').getValue();
  
//   if (!siteType || !keyword) {
//     console.log('サイト種別またはキーワードが入力されていません');
//     return;
//   }
  
//   // D2からジャンルを取得
//   var genre = sheet.getRange('D2').getValue();
  
//   // カテゴリIDを取得
//   var categoryId = getCategoryId_(genre, siteType);
  
//   // キーワードをURLエンコード
//   var encodedKeyword = encodeURIComponent(keyword);
  
//   var url = '';
  
//   if (siteType === '楽天') {
//     // 楽天のURL生成
//     if (categoryId) {
//       url = 'https://search.rakuten.co.jp/search/mall/' + encodedKeyword + '/' + categoryId + '/?s=2';
//     } else {
//       url = 'https://search.rakuten.co.jp/search/mall/' + encodedKeyword + '/?s=2';
//     }
//   } else if (siteType === 'SBA') {
//     // SBAのURL生成
//     var categoryParam = categoryId ? categoryId : '';
//     var periodParam = '';
    
//     // 日付をYYYY-MM-DD形式に変換
//     if (period) {
//       if (period instanceof Date) {
//         // Dateオブジェクトの場合
//         var year = period.getFullYear();
//         var month = ('0' + (period.getMonth() + 1)).slice(-2);
//         var day = ('0' + period.getDate()).slice(-2);
//         periodParam = year + '-' + month + '-' + day;
//         console.log('日付をフォーマット:', periodParam);
//       } else {
//         // 文字列の場合（YYYY-MM-DD形式を期待）
//         periodParam = String(period);
//         // もし文字列が日付っぽくない場合は、パースを試みる
//         if (periodParam.indexOf('GMT') > -1 || periodParam.indexOf('日本標準時') > -1) {
//           var dateObj = new Date(periodParam);
//           if (!isNaN(dateObj.getTime())) {
//             var year = dateObj.getFullYear();
//             var month = ('0' + (dateObj.getMonth() + 1)).slice(-2);
//             var day = ('0' + dateObj.getDate()).slice(-2);
//             periodParam = year + '-' + month + '-' + day;
//             console.log('日付文字列をフォーマット:', periodParam);
//           }
//         }
//       }
//     }
    
//     url = 'https://www.starbuyers-global-auction.com/market_price?sort=exhibit_date&direction=desc&limit=30' +
//           '&item_category=' + categoryParam +
//           '&keywords=' + encodedKeyword +
//           '&rank=&price_from=&price_to=' +
//           '&exhibit_date_from=' + periodParam +
//           '&exhibit_date_to=&accessory=';
//   } else {
//     console.log('サイト種別は「楽天」または「SBA」を指定してください');
//     return;
//   }
  
//   // B28にURLを設定
//   sheet.getRange('B28').setValue(url);
//   console.log('B28に検索URLを設定しました:', url);
// }

// // 卸先シートからカテゴリIDを取得
// function getCategoryId_(genre, siteType) {
//   if (!genre) return '';
  
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var wholesaleSheet = ss.getSheetByName('卸先');
  
//   if (!wholesaleSheet) {
//     console.log('「卸先」シートが見つかりません');
//     return '';
//   }
  
//   // A列から最終行を取得
//   var lastRow = wholesaleSheet.getLastRow();
//   if (lastRow < 2) return ''; // ヘッダー行のみの場合
  
//   // C列（ジャンル）の値を取得
//   var genreRange = wholesaleSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  
//   // ジャンルが一致する行を探す
//   var targetRow = -1;
//   for (var i = 0; i < genreRange.length; i++) {
//     if (genreRange[i][0] === genre) {
//       targetRow = i + 2; // 実際の行番号（ヘッダー行を考慮）
//       break;
//     }
//   }
  
//   if (targetRow === -1) {
//     console.log('ジャンル「' + genre + '」が卸先シートに見つかりません');
//     return '';
//   }
  
//   // カテゴリIDを取得（I列:SBA、J列:楽天）
//   var categoryId = '';
//   if (siteType === 'SBA') {
//     categoryId = wholesaleSheet.getRange(targetRow, 9).getValue(); // I列
//   } else if (siteType === '楽天') {
//     categoryId = wholesaleSheet.getRange(targetRow, 10).getValue(); // J列
//   }
  
//   return String(categoryId || '');
// }

// /** ============== 認証情報管理（推奨） ============== **/
// // スクリプトプロパティに認証情報を保存する（初回設定用）
// function setStarBuyersCredentials() {
//   // この関数を一度実行して認証情報を保存してください
//   var scriptProperties = PropertiesService.getScriptProperties();
//   scriptProperties.setProperty('STARBUYERS_EMAIL', 'your_email@example.com');
//   scriptProperties.setProperty('STARBUYERS_PASSWORD', 'your_password');
// }

// // スクリプトプロパティから認証情報を取得
// function getStarBuyersCredentials_() {
//   var scriptProperties = PropertiesService.getScriptProperties();
//   var email = scriptProperties.getProperty('STARBUYERS_EMAIL');
//   var password = scriptProperties.getProperty('STARBUYERS_PASSWORD');
  
//   // スクリプトプロパティに設定されていない場合はハードコードされた値を使用（非推奨）
//   if (!email || !password) {
//     console.warn('認証情報がスクリプトプロパティに設定されていません。setStarBuyersCredentials()を実行してください。');
//     return {
//       email: 'inui.hur@gmail.com',
//       password: 'hur22721'
//     };
//   }
  
//   return { email: email, password: password };
// }

// /** ============== メイン ============== **/
// // B28のHTML（楽天/StarBuyersの場合はURL可） → 自動判定 → パース → B10〜B17に出力（B17は引用サイト）
// // 使い方：
// //   1. HTMLソースをB28以降に貼り付け、または
// //   2. 楽天のURLをB28に貼り付け（自動でHTMLを取得）、または
// //   3. Star BuyersのURLをB28に貼り付け（ログインして自動でHTMLを取得）
// function fillTemplateFromHtml_B10_V19(sheetName) {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
//   if (!sh) throw new Error('対象シートが見つかりません');

//   // B28からHTML取得（楽天/StarBuyers URLの場合は自動フェッチ）
//   var html = readHtmlFromB28_(sh);
//   var src = detectSource_(html);

//   var items = (src === 'starbuyers') ? parseStarBuyersFromHtml_(html)
//             : (src === 'aucfan')     ? parseAucfanFromHtml_(html)
//             : (src === 'rakuten')    ? parseRakutenFromHtml_(html)
//             :                          parseStarBuyersFromHtml_(html);

//   if (items.length === 0) {
//     // フォールバック：他のパーサーも試す
//     if (src !== 'aucfan') {
//       items = parseAucfanFromHtml_(html);
//     }
//     if (items.length === 0 && src !== 'rakuten') {
//       items = parseRakutenFromHtml_(html);
//     }
//   }

//   // 出力件数制限（C〜V）
//   if (items.length > MAX_OUTPUT_ITEMS) items = items.slice(0, MAX_OUTPUT_ITEMS);

//   writeItemsToTemplate_B10_V19_(sh, items);
  
//   // D2の査定ジャンルに基づいてB16（服サイズ）を自動設定
//   handleClothingSizeField_(sh);
// }

// /** 28行目以降をB列含めて全クリア */
// function clearHtmlFromB28Down(sheetName) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
//   if (!sh) throw new Error('対象シートが見つかりません');
//   const last = sh.getLastRow();
//   const lastCol = sh.getLastColumn();
//   if (last >= ROW_HTML_START) {
//     sh.getRange(ROW_HTML_START, 1, last - ROW_HTML_START + 1, lastCol).clearContent();
//   }
// }

// /**
//  * 10行目・11行目・14〜19行目・27行目以降の「C列以降」をクリア
//  * （行ヘッダBは残す）
//  */
// function clearSpecificRows() {
//   var sh = SpreadsheetApp.getActiveSheet();
//   var last = sh.getLastRow();
//   var lastCol = sh.getLastColumn();
//   var widthFromC = Math.max(0, lastCol - (COL_C - 1)); // C〜最終列の幅

//   if (widthFromC <= 0) return;

//   // 10行目
//   if (last >= ROW_CLEAR_10) {
//     sh.getRange(ROW_CLEAR_10, COL_C, 1, widthFromC).clearContent();
//   }
//   // 11行目
//   if (last >= ROW_CLEAR_11) {
//     sh.getRange(ROW_CLEAR_11, COL_C, 1, widthFromC).clearContent();
//   }
//   // 14〜19行目
//   if (last >= ROW_CLEAR_BLOCK_FROM) {
//     var to = Math.min(ROW_CLEAR_BLOCK_TO, last);
//     var rowCount = to - ROW_CLEAR_BLOCK_FROM + 1;
//     if (rowCount > 0) {
//       sh.getRange(ROW_CLEAR_BLOCK_FROM, COL_C, rowCount, widthFromC).clearContent();
//     }
//   }
//   clearHtmlFromB28Down("")
// }
