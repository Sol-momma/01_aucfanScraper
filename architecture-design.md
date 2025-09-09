# オークファン商品ランク機能 アーキテクチャ設計書

## 📋 タスク概要
オークファンから取得した商品データに、商品の状態（コンディション）を解析してランク（S、SA、A、B、C等）に変換する機能を追加する。

## 🎯 目標
スクレイピング → 商品の状態を取得 → S/A/B/Cに変換 → スプレッドシートに記載

## 🏗️ 現在の状況分析

### 既存のコード構造
- `01_aucfanScraper.js` にオークファンのスクレイピング機能が実装済み
- `parseAucfanFromHtml_()` 関数で商品情報を抽出中
- 現在は `rank: ""` で空文字のまま

### 対象サイト
1. **ヤフオク**: "商品状態 中古" / "商品状態 新品"
2. **メルカリ**: "目立った傷や汚れなし" / "やや傷や汚れあり"

## 📐 アーキテクチャ設計

### 1. データ抽出レイヤー（Extract Layer）
```javascript
/**
 * 商品状態のテキストを抽出する関数
 * @param {string} htmlBlock - 商品ブロックのHTML
 * @return {string} 商品状態のテキスト（例: "中古", "目立った傷や汚れなし"）
 */
function extractConditionText_(htmlBlock)

/**
 * サイト種別を判定する関数  
 * @param {string} htmlBlock - 商品ブロックのHTML
 * @return {string} サイト名（"yahoo" または "mercari"）
 */
function detectSiteType_(htmlBlock)
```

### 2. 変換レイヤー（Transform Layer）
```javascript
/**
 * ヤフオクの商品状態をランクに変換
 * @param {string} conditionText - 商品状態テキスト
 * @return {string} ランク（S, SA, A, B, C, D）
 */
function convertYahooConditionToRank_(conditionText)

/**
 * メルカリの商品状態をランクに変換
 * @param {string} conditionText - 商品状態テキスト  
 * @return {string} ランク（S, SA, A, B, C, D）
 */
function convertMercariConditionToRank_(conditionText)

/**
 * 統一ランク変換関数（メイン）
 * @param {string} conditionText - 商品状態テキスト
 * @param {string} siteType - サイト種別
 * @return {string} ランク
 */
function convertConditionToRank_(conditionText, siteType)
```

### 3. マッピング定義レイヤー（Mapping Layer）
```javascript
/**
 * ヤフオク商品状態とランクのマッピング定義
 */
const YAHOO_CONDITION_RANK_MAP = {
  '新品': 'S',
  '未使用': 'S', 
  '未使用に近い': 'SA',
  '目立った傷や汚れなし': 'A',
  'やや傷や汚れあり': 'B',
  '傷や汚れあり': 'C',
  '中古': 'B',  // デフォルト
  '全体的に状態が悪い': 'D'
}

/**
 * メルカリ商品状態とランクのマッピング定義
 */
const MERCARI_CONDITION_RANK_MAP = {
  '新品、未使用': 'S',
  '未使用に近い': 'SA',
  '目立った傷や汚れなし': 'A', 
  'やや傷や汚れあり': 'B',
  '傷や汚れあり': 'C',
  '全体的に状態が悪い': 'D'
}
```

## 🔧 実装手順

### Phase 1: 商品状態抽出機能の実装
1. `extractConditionText_()` 関数を実装
   - ヤフオクの「商品状態」パターンを検索
   - メルカリの状態説明パターンを検索
   - 正規表現で状態テキストを抽出

### Phase 2: サイト判定機能の実装  
2. `detectSiteType_()` 関数を実装
   - HTMLブロック内の特徴的な要素でサイトを判定
   - ヤフオク: "aucfan.com" + Yahoo関連のクラス名
   - メルカリ: "mercari" 関連のパターン

### Phase 3: ランク変換機能の実装
3. マッピング定義オブジェクトを作成
4. `convertYahooConditionToRank_()` 関数を実装
5. `convertMercariConditionToRank_()` 関数を実装  
6. `convertConditionToRank_()` 統一関数を実装

### Phase 4: 既存コードとの統合
7. `parseAucfanFromHtml_()` 関数を修正
   - 商品状態抽出を追加
   - ランク変換を追加
   - `rank: ""` を `rank: convertedRank` に変更

### Phase 5: テスト・デバッグ
8. `testAucfanRankExtraction()` テスト関数を作成
9. 実データでのテスト実行
10. マッピングの調整・改善

## 🎨 実装例

### 修正後の parseAucfanFromHtml_ 関数の流れ
```javascript
function parseAucfanFromHtml_(html) {
  // ... 既存のコード ...
  
  while ((m = re.exec(html)) !== null) {
    const block = m[1];
    
    // 既存の抽出処理
    // ...
    
    // 🆕 商品状態抽出とランク変換
    const conditionText = extractConditionText_(block);
    const siteType = detectSiteType_(block);  
    const rank = convertConditionToRank_(conditionText, siteType);
    
    if (detailUrl || imageUrl || price || endTxt || title) {
      items.push(
        createItemData_({
          title: title,
          detailUrl: detailUrl,
          imageUrl: imageUrl,
          date: endTxt || "",
          rank: rank, // 🆕 ここに変換されたランクが入る
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
```

## 🧪 テスト計画

### テストケース
1. **ヤフオクの商品状態テスト**
   - "商品状態 新品" → "S"
   - "商品状態 中古" → "B"
   - "商品状態 未使用に近い" → "SA"

2. **メルカリの商品状態テスト**  
   - "目立った傷や汚れなし" → "A"
   - "やや傷や汚れあり" → "B"

3. **エラーケーステスト**
   - 商品状態が見つからない場合 → ""（空文字）
   - 未知の状態テキスト → "B"（デフォルト）

## 🔄 今後の拡張性

### 容易に調整可能な設計
- マッピング定義は独立したオブジェクト
- 新しいサイト追加時は判定関数と変換関数を追加するだけ
- ランク基準の変更は マッピングオブジェクトを修正するだけ

### 将来的な改善案
- 機械学習による自動ランク判定
- ユーザー設定可能なカスタムマッピング
- ランク判定の信頼度スコア追加

## 📝 開発開始時のチェックリスト

- [ ] Phase 1: `extractConditionText_()` 実装
- [ ] Phase 2: `detectSiteType_()` 実装  
- [ ] Phase 3: マッピング定義 + 変換関数実装
- [ ] Phase 4: `parseAucfanFromHtml_()` 修正
- [ ] Phase 5: テスト関数作成 + 動作確認

## 🎓 初心者向け解説

### なぜこの設計にしたのか？

#### 1. **レイヤー分離の思想**
コードを3つの責任に分けました：
- **抽出**: HTMLから状態テキストを見つける
- **変換**: 状態テキストをランクに変換  
- **マッピング**: どの状態がどのランクかの対応表

**理由**: 後で調整しやすくなるから！例えば、ランク基準を変えたい時は、マッピング部分だけ修正すればOK。

#### 2. **関数名の命名規則**  
- `extract〜()`: HTMLから情報を取り出す
- `convert〜()`: データを変換する
- `detect〜()`: 何かを判定する

**理由**: 関数名を見ただけで何をするかわかるので、バグが起きにくく、他の人も理解しやすい。

#### 3. **段階的開発（Phase分け）**
いきなり全部作らず、5段階に分けました。

**理由**: 各段階でテストできるので、どこで問題が起きたかすぐわかる。安全な開発ができます。

### 開発の進め方

**Phase 1の`extractConditionText_()`関数から始めるのがベスト！**

理由：
- 一番基礎的な部分
- この関数ができれば、残りは組み合わせるだけ
- テストしやすい（HTMLを渡せば結果がすぐわかる）

---
*この設計書は開発進捗に応じて随時更新していきます*
