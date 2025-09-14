/**
 * データ書き込み処理を担当するモジュール
 * スプレッドシートへの商品データ書き込みを共通化
 */

// ===== 定数定義 =====
// 設定値は00_configManager.gsから動的に取得
// フォールバック値は00_main.gsの定数を使用

// ===== 動的設定値取得関数 =====

/**
 * 出力設定値を取得（設定シートまたはフォールバック値）
 */
function getOutputSettings_() {
  try {
    return {
      maxItems: getConfig_("OUTPUT_MAX_ITEMS") || 20,
      maxCols: getConfig_("OUTPUT_MAX_COLS") || 20,
      startCol: getConfig_("OUTPUT_START_COL") || 3,
      applyClipWrap: getConfig_("OUTPUT_CLIP_WRAP") !== false,
    };
  } catch (e) {
    console.warn("設定値取得エラー、フォールバック値を使用:", e.message);
    return {
      maxItems: 20,
      maxCols: 20,
      startCol: 3,
      applyClipWrap: true,
    };
  }
}

// ===== 共通ユーティリティ関数 =====

/**
 * ランクを正規化（データ書き込み用）
 */
function normalizeRankForWrite_(rank) {
  if (!rank) return "";
  var normalized = rank.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
  });
  var validRanks = ["N", "S", "SA", "A", "AB", "B", "BC", "C"];
  return validRanks.includes(normalized.toUpperCase())
    ? normalized.toUpperCase()
    : "";
}

// ===== メインの書き込み関数 =====
/**
 * 指定されたセクションにアイテムデータを配置
 */
function writeItemsToSpecificSection_(sheet, items, startRow, sectionName) {
  console.log(
    sectionName +
      "セクションにデータ配置開始: " +
      (items ? items.length : 0) +
      "件"
  );

  // 楽天の場合は受け取った商品データを詳細確認
  if (sectionName === "楽天" && items && items.length > 0) {
    console.log("=== writeItemsToSpecificSection_楽天データ受信確認 ===");
    items.slice(0, 3).forEach(function (item, index) {
      console.log("受信商品" + (index + 1) + ":", {
        title: item.title ? item.title.substring(0, 30) + "..." : "なし",
        soldOut: item.soldOut,
        soldOutプロパティ存在: item.hasOwnProperty("soldOut"),
        全プロパティ: Object.keys(item),
      });
    });
  }

  if (!items || items.length === 0) {
    console.log(sectionName + "のデータがありません");
    return;
  }

  // スプレッドシートのレイアウトに合わせた行マッピング
  // セクションごとに異なるレイアウトに対応
  var rowMapping;

  if (sectionName === "楽天") {
    // 楽天セクション: 31行目開始、32行目が画像URL
    rowMapping = {
      詳細URL: startRow + 0, // 31行目
      画像URL: startRow + 1, // 32行目
      商品名: startRow + 4, // 35行目
      日付: startRow + 5, // 36行目
      ランク: startRow + 6, // 37行目
      服サイズ: startRow + 7, // 38行目
      カラー: startRow + 8, // 39行目
      価格: startRow + 9, // 40行目
      売り切れ: startRow + 10, // 41行目
    };
  } else if (sectionName === "エコリング") {
    // エコリングセクション: 18行目開始、19行目が画像URL
    rowMapping = {
      詳細URL: startRow + 0, // 18行目
      画像URL: startRow + 1, // 19行目
      商品名: startRow + 4, // 22行目
      日付: startRow + 5, // 23行目
      ランク: startRow + 6, // 24行目
      服サイズ: startRow + 7, // 25行目
      カラー: startRow + 8, // 26行目
      価格: startRow + 9, // 27行目
    };
  } else if (sectionName === "オークファン") {
    // オークファンセクション: 実際のレイアウトに合わせた固定行番号
    rowMapping = {
      詳細URL: 44, // 詳細URL: 44行目
      画像URL: 45, // 画像URL: 45行目
      画像: 46, // 画像: 46行目
      計算に入れるか: 47, // 計算に入れるか: 47行目
      商品名: 48, // 商品名: 48行目
      日付: 49, // 日付: 49行目
      ランク: 50, // ランク: 50行目
      服サイズ: 51, // 服サイズ: 51行目
      カラー: 52, // カラー: 52行目
      価格: 53, // 価格: 53行目
    };
  } else if (sectionName === "ヤフオク") {
    // ヤフオクセクション: 実際のレイアウトに合わせた固定行番号
    rowMapping = {
      詳細URL: 45, // 詳細URL: 45行目
      画像URL: 46, // 画像URL: 46行目
      画像: 47, // 画像: 47行目
      計算に入れるか: 48, // 計算に入れるか: 48行目
      商品名: 49, // 商品名: 49行目
      日付: 50, // 日付: 50行目
      ランク: 51, // ランク: 63行目
      服サイズ: 52, // 服サイズ: 52行目
      カラー: 53, // カラー: 53行目
      価格: 54, // 価格: 54行目
    };
  } else {
    // その他のセクション: 従来の配置
    rowMapping = {
      詳細URL: startRow + 0, // 開始行
      画像URL: startRow + 1, // 開始行+1
      商品名: startRow + 4, // 開始行+4
      日付: startRow + 5, // 開始行+5
      ランク: startRow + 6, // 開始行+6
      服サイズ: startRow + 7, // 開始行+7
      カラー: startRow + 8, // 開始行+8
      価格: startRow + 9, // 開始行+9
    };
  }

  // 出力件数制限
  var limitedItems = items.slice(0, MAX_OUTPUT_ITEMS);

  // 楽天セクション専用のデバッグ（強制売り切れ設定は無効化）
  if (sectionName === "楽天") {
    console.log("=== 楽天売り切れ情報デバッグ ===");
    console.log("limitedItems数:", limitedItems.length);

    // 最初の3商品の売り切れ状況を確認
    limitedItems.slice(0, 3).forEach(function (item, index) {
      console.log("商品" + (index + 1) + ":");
      console.log("  title:", (item.title || "").substring(0, 30) + "...");
      console.log("  soldOut:", item.soldOut);
      console.log("  soldOutプロパティ存在:", item.hasOwnProperty("soldOut"));
      console.log("  全プロパティ:", Object.keys(item));
    });

    // テスト用：最初の3商品を強制的に売り切れに設定（無効化）
    // console.log("★★★ テスト：最初の3商品を強制的に売り切れに設定 ★★★");
    // for (var i = 0; i < Math.min(3, limitedItems.length); i++) {
    //   limitedItems[i].soldOut = true;
    //   console.log(
    //     "商品" + (i + 1) + "を売り切れに設定完了:",
    //     limitedItems[i].soldOut
    //   );
    // }
  }

  // データ配列を準備
  var rowArrays = {
    詳細URL: limitedItems.map((it) => it.detailUrl || ""),
    画像URL: limitedItems.map((it) => it.imageUrl || ""),
    画像: limitedItems.map(() => ""), // 画像行は空のまま（画像表示用の数式が入る）
    計算に入れるか: limitedItems.map(() => ""), // 空のまま（手動入力用）
    商品名: limitedItems.map((it) => it.title || ""),
    日付: limitedItems.map((it) => it.date || ""),
    ランク: limitedItems.map((it) => normalizeRankForWrite_(it.rank || "")),
    服サイズ: limitedItems.map(() => ""), // 空のまま（手動入力用）
    カラー: limitedItems.map(() => ""), // 空のまま（手動入力用）
    引用サイト: limitedItems.map((it) => it.source || ""),
    価格: limitedItems.map((it) => it.price || ""),
    ショップ: limitedItems.map((it) => it.shop || ""),
    売り切れ: limitedItems.map((it) => it.soldOut === true), // チェックボックス用のboolean値
  };

  // 楽天セクション専用：売り切れ配列の確認
  if (sectionName === "楽天") {
    console.log("売り切れ配列:", rowArrays["売り切れ"]);
    console.log("rowMapping売り切れ行:", rowMapping["売り切れ"]);
  }

  // 各ラベルに対応する行にデータを配置
  Object.keys(rowMapping).forEach(function (label) {
    var targetRow = rowMapping[label];
    var data = rowArrays[label];

    // 画像行はスキップ（画像表示用の数式が入るため、データは書き込まない）
    if (label === "画像") {
      console.log(
        sectionName +
          "の「" +
          label +
          "」行はスキップしました（数式保持のため）"
      );
      return;
    }

    // オークファンの場合は、空のデータでも行を処理する（レイアウト維持のため）
    var shouldProcess =
      (data && data.length > 0) || sectionName === "オークファン";

    if (shouldProcess) {
      // 売り切れ行の場合は詳細ログ
      if (label === "売り切れ" && sectionName === "楽天") {
        console.log("=== 楽天売り切れ行処理詳細 ===");
        console.log("targetRow:", targetRow);
        console.log("data:", data);
        console.log("data.length:", data ? data.length : "なし");
        console.log("data内容詳細:", data.slice(0, 5));
        console.log(
          "データ型:",
          data.slice(0, 5).map((d) => typeof d)
        );
        console.log("trueの数:", data.filter((d) => d === true).length);
        console.log("falseの数:", data.filter((d) => d === false).length);
      }

      // データ配列を準備
      var values = [];
      for (var i = 0; i < MAX_OUTPUT_ITEMS; i++) {
        var value = data && i < data.length ? data[i] : "";

        // ランクの場合、有効な値のみ設定
        if (label === "ランク") {
          // 有効なランク値のみ設定、空の場合は空文字列
          var validRanks = ["N", "S", "SA", "A", "AB", "B", "BC", "C"];
          if (value && validRanks.indexOf(value.toUpperCase()) > -1) {
            values.push(value.toUpperCase());
          } else {
            values.push("");
          }
        } else {
          values.push(value);
        }
      }

      // 該当行をクリア（特別処理が必要な行を判定）
      // 統一クリア関数を使用
      // 注意：最初に一括消去処理が実行されているため、個別クリアは不要
      // 売り切れ行とチェックボックス系は事前クリアをスキップ（一括書き込みで上書き）
      // clearRowByItemType_(sheet, label, targetRow);

      // データを設定（有効なデータがある場合のみ、オークファンは常に処理）
      var hasValidData =
        values.some((v) => v !== "" && v !== null && v !== undefined) ||
        sectionName === "オークファン";

      if (hasValidData) {
        try {
          var range = sheet.getRange(targetRow, COL_C, 1, MAX_OUTPUT_ITEMS);

          // ランクの場合は特別な処理
          if (label === "ランク") {
            console.log(
              sectionName + "のランクデータ:",
              values.filter((v) => v !== "").join(", ")
            );

            // ランクNが含まれているかチェック
            var rankNCount = values.filter((v) => v === "N").length;
            if (rankNCount > 0) {
              console.log(
                sectionName + "にランクNが" + rankNCount + "件含まれています"
              );
            }

            // ランクの場合は個別にセルに値を設定
            for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
              var value = values[col];
              var cellRange = sheet.getRange(targetRow, COL_C + col);
              cellRange.setValue(value);
            }
          } else if (label === "画像URL") {
            // 画像URL行の場合は数式を保持してデータを設定
            console.log(
              sectionName + "の画像URLデータ:",
              values.filter((v) => v !== "").slice(0, 3)
            );

            // 一括処理のため、まず範囲全体の数式を取得
            var range = sheet.getRange(targetRow, COL_C, 1, MAX_OUTPUT_ITEMS);
            var formulas = range.getFormulas()[0]; // 1行分の数式配列
            var newValues = [[]]; // 2次元配列として初期化

            for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
              var formula = formulas[col];
              var value = values[col];

              if (formula) {
                // 数式がある場合は保持（既存の数式をそのまま使用）
                newValues[0][col] = formula;
                console.log(
                  "数式を保持: " +
                    formula +
                    " (行" +
                    targetRow +
                    "列" +
                    (COL_C + col) +
                    ")"
                );
              } else if (value && value !== "") {
                // 数式がなく、データがある場合のみ設定
                newValues[0][col] = value;
                console.log(
                  "画像URL設定:",
                  value,
                  "(行" + targetRow + "列" + (COL_C + col) + ")"
                );
              } else {
                // 数式もデータもない場合は空文字
                newValues[0][col] = "";
              }
            }

            // 一括で値を設定
            range.setValues(newValues);
          } else if (label === "服サイズ" || label === "カラー") {
            // 服サイズとカラーの場合は背景色も保持
            console.log(
              sectionName + "の" + label + "データ（背景色保持）:",
              values.filter((v) => v !== "").slice(0, 3)
            );

            for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
              var cellRange = sheet.getRange(targetRow, COL_C + col);
              var validation = cellRange.getDataValidation();
              var bgColor = cellRange.getBackground();
              var value = values[col];

              // 値を設定（空の場合も含む）
              if (value !== null && value !== undefined) {
                cellRange.setValue(value);
              }

              // データ検証を復元
              if (validation) {
                cellRange.setDataValidation(validation);
              }

              // 背景色を復元
              if (bgColor) {
                cellRange.setBackground(bgColor);
              }
            }
          } else if (label === "売り切れ" || label === "計算に入れるか") {
            // チェックボックス系は一括書き込み（データ入力規則を保持）
            console.log(
              sectionName + "の" + label + "データ（チェックボックス）:",
              values.slice(0, 5)
            );

            // 一括で値を設定（データ入力規則は自動保持される）
            // 事前クリアなしで直接上書きすることで高速化
            range.setValues([values]);
          } else {
            // その他の場合は従来の処理
            // データ入力規則と背景色を一時的に保存
            var validationRules = [];
            var bgColors = [];

            for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
              var cellRange = sheet.getRange(targetRow, COL_C + col);
              var validation = cellRange.getDataValidation();
              var bgColor = cellRange.getBackground();
              validationRules.push(validation);
              bgColors.push(bgColor);
              if (validation) {
                cellRange.clearDataValidations();
              }
            }

            // データを設定
            range.setValues([values]);

            // データ入力規則と背景色を復元
            for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
              var cellRange = sheet.getRange(targetRow, COL_C + col);
              if (validationRules[col]) {
                cellRange.setDataValidation(validationRules[col]);
              }
              if (bgColors[col]) {
                cellRange.setBackground(bgColors[col]);
              }
            }
          }

          // 折返し抑止
          if (APPLY_CLIP_WRAP) {
            var outLen = Math.min(limitedItems.length, MAX_OUTPUT_COLS);
            if (outLen > 0) {
              sheet
                .getRange(targetRow, COL_C, 1, outLen)
                .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
            }
          }
        } catch (e) {
          console.error(
            sectionName + "の「" + label + "」書き込みエラー:",
            e.message
          );
        }
      } else if (label === "ランク") {
        console.log(sectionName + "のランクデータは空のためスキップしました");
      }
    }
  });

  console.log(
    sectionName +
      "セクション（行" +
      startRow +
      "～" +
      (startRow + 12) +
      "）にデータを配置しました"
  );

  console.log(
    sectionName + "セクションにデータ配置完了: " + limitedItems.length + "件"
  );
}

// ===== シンプル統一クリア関数 =====

/**
 * 行をシンプルにクリア（値のみクリア、書式は保持）
 * @param {Sheet} sheet シート
 * @param {number} row 行番号
 * @param {boolean} preserveFormulas 数式を保持するかどうか（画像行用）
 */
function clearRowSimple_(sheet, row, preserveFormulas) {
  try {
    var settings = getOutputSettings_();

    var range = sheet.getRange(row, settings.startCol, 1, settings.maxCols);

    if (preserveFormulas) {
      // 数式保持クリア（高速一括処理）
      var formulas = range.getFormulas()[0];
      var values = range.getValues()[0];
      var newValues = [];

      for (var col = 0; col < settings.maxCols; col++) {
        if (formulas[col]) {
          // 数式がある場合は現在の値を保持
          newValues.push(values[col]);
        } else {
          // 数式がない場合はクリア
          newValues.push("");
        }
      }

      // 一括で値を設定
      range.setValues([newValues]);
    } else {
      // 通常クリア（値のみクリア、書式は自動保持）
      var emptyValues = new Array(settings.maxCols).fill("");
      range.setValues([emptyValues]);
    }

    console.log(
      "行" + row + "をクリアしました" + (preserveFormulas ? "（数式保持）" : "")
    );
  } catch (e) {
    console.error("行" + row + "のクリアエラー:", e.message);
  }
}

/**
 * 項目の特性に応じて適切な方法でクリア（シンプル版）
 * @param {Sheet} sheet シート
 * @param {string} itemType 項目タイプ
 * @param {number} row 行番号
 */
function clearRowByItemType_(sheet, itemType, row) {
  try {
    // 売り切れ行とチェックボックス系は事前クリアをスキップ（後で一括上書きするため）
    // ただし、リセット時は強制的にクリアする
    var skipClear = itemType === "売り切れ" || itemType === "計算に入れるか";

    if (skipClear) {
      // チェックボックス系の行は強制的にfalseに設定してクリア
      var settings = getOutputSettings_();
      var checkboxRange = sheet.getRange(
        row,
        settings.startCol,
        1,
        settings.maxCols
      );
      var checkboxValues = checkboxRange.getValues()[0];
      for (var col = 0; col < settings.maxCols; col++) {
        if (typeof checkboxValues[col] === "boolean") {
          checkboxValues[col] = false;
        } else {
          checkboxValues[col] = "";
        }
      }
      checkboxRange.setValues([checkboxValues]);
      console.log(
        itemType + " (行" + row + ") のチェックボックスをリセットしました"
      );
      return;
    }

    // 画像行のみ数式を保持、その他は全て値のみクリア
    var preserveFormulas = itemType === "画像" || itemType === "画像URL";
    clearRowSimple_(sheet, row, preserveFormulas);

    console.log(itemType + " (行" + row + ") をクリアしました");
  } catch (e) {
    console.error(itemType + " (行" + row + ") のクリアエラー:", e.message);
  }
}

/**
 * セクション全体を一括でクリア（高速一括処理版）
 * @param {Sheet} sheet シート
 * @param {Object} rowMapping 行マッピング
 */
function clearSectionBatch_(sheet, rowMapping) {
  try {
    var settings = getOutputSettings_();

    // 各行を個別に処理（対象行のみ）
    Object.keys(rowMapping).forEach(function (itemType) {
      var row = rowMapping[itemType];
      var preserveFormulas = itemType === "画像" || itemType === "画像URL";

      // 売り切れ行とチェックボックス系は事前クリアをスキップ（後で一括上書きするため）
      // ただし、リセット時は強制的にクリアする
      var skipClear = itemType === "売り切れ" || itemType === "計算に入れるか";

      if (skipClear) {
        // チェックボックス系の行は強制的にfalseに設定してクリア
        var checkboxRange = sheet.getRange(
          row,
          settings.startCol,
          1,
          settings.maxCols
        );
        var checkboxValues = checkboxRange.getValues()[0];
        for (var col = 0; col < settings.maxCols; col++) {
          if (typeof checkboxValues[col] === "boolean") {
            checkboxValues[col] = false;
          } else {
            checkboxValues[col] = "";
          }
        }
        checkboxRange.setValues([checkboxValues]);
        console.log(
          itemType + " (行" + row + ") のチェックボックスをリセットしました"
        );
        return;
      }

      // 行の範囲を取得
      var range = sheet.getRange(row, settings.startCol, 1, settings.maxCols);

      if (preserveFormulas) {
        // 数式保持クリア（高速一括処理）
        var formulas = range.getFormulas()[0];
        var values = range.getValues()[0];
        var newValues = [];

        for (var col = 0; col < settings.maxCols; col++) {
          if (formulas[col]) {
            // 数式がある場合は現在の値を保持
            newValues.push(values[col]);
          } else {
            // 数式がない場合はクリア
            newValues.push("");
          }
        }

        // 一括で値を設定
        range.setValues([newValues]);
      } else {
        // 通常クリア（一括処理）
        var emptyValues = new Array(settings.maxCols).fill("");
        range.setValues([emptyValues]);
      }

      console.log(itemType + " (行" + row + ") をクリアしました");
    });

    console.log(
      "セクションクリア完了（" + Object.keys(rowMapping).length + "項目）"
    );
  } catch (e) {
    console.error("セクションクリアエラー:", e.message);
    // エラーの場合は従来の方法にフォールバック
    Object.keys(rowMapping).forEach(function (itemType) {
      var row = rowMapping[itemType];
      clearRowByItemType_(sheet, itemType, row);
    });
  }
}

/**
 * セクション全体をシンプルにクリア
 * @param {Sheet} sheet シート
 * @param {string} sectionName セクション名
 */
function clearSectionForDataFetch_(sheet, sectionName) {
  try {
    console.log(sectionName + "セクションのクリア開始");

    // 詳細URLの行位置を取得
    var detailUrlRow = getDetailUrlRow_(sectionName);
    if (!detailUrlRow) {
      console.error(sectionName + "の詳細URL行が見つかりません");
      return;
    }

    // 各項目の行を計算（詳細URLを基準とした相対位置）
    var rowMapping = {
      詳細URL: detailUrlRow + 0, // 基準行
      画像URL: detailUrlRow + 1, // +1
      画像: detailUrlRow + 2, // +2（数式保持）
      計算に入れるか: detailUrlRow + 3, // +3
      商品名: detailUrlRow + 4, // +4
      日付: detailUrlRow + 5, // +5
      ランク: detailUrlRow + 6, // +6
      服サイズ: detailUrlRow + 7, // +7
      カラー: detailUrlRow + 8, // +8
      価格: detailUrlRow + 9, // +9
    };

    // 楽天の場合は売り切れ行も追加
    if (sectionName === "楽天") {
      rowMapping["売り切れ"] = detailUrlRow + 10; // +10（価格の次の行）
      console.log(
        "楽天セクション：売り切れ行もクリア対象に追加（行" +
          rowMapping["売り切れ"] +
          "）"
      );
    }

    // 一括クリア処理（高速化）
    clearSectionBatch_(sheet, rowMapping);

    console.log(sectionName + "セクションのクリア完了");
  } catch (e) {
    console.error(sectionName + "セクションのクリアエラー:", e.message);
  }
}

// ===== セクション別データクリア関数 =====

/**
 * セクション名から詳細URLの行位置を取得
 * @param {string} sectionName セクション名
 * @return {number|null} 詳細URLの行番号、見つからない場合はnull
 */
function getDetailUrlRow_(sectionName) {
  var detailUrlRows = {
    SBA: 5,
    エコリング: 18,
    楽天: 31,
    オークファン: 44,
    ヤフオク: 45,
  };

  return detailUrlRows[sectionName] || null;
}

/**
 * 全サイトのデータを一括クリア（対象行のみ処理版）
 */
function clearAllSitesDataFast_(sheet) {
  try {
    console.log("全サイト一括クリア開始");

    // 全サイトを順次クリア（対象行のみ処理）
    // オークファンは別シートで管理するため除外
    var siteNames = ["SBA", "エコリング", "楽天", "ヤフオク"];

    siteNames.forEach(function (siteName) {
      clearSectionForDataFetch_(sheet, siteName);
    });

    console.log(
      "注意: オークファンは別シートで管理されるため、全サイトクリアから除外されました"
    );

    console.log("全サイト一括クリア完了");
  } catch (e) {
    console.error("全サイト一括クリアエラー:", e.message);
  }
}

// ===== 衣類サイズ処理関数 =====

/**
 * 全セクションの衣類サイズフィールドを処理
 */
function handleClothingSizeFieldForAllSections_(sheet) {
  var sections = [
    { name: "SBA", startRow: getSiteConfig_("sba").startRow },
    { name: "エコリング", startRow: getSiteConfig_("ecoring").startRow },
    { name: "楽天", startRow: getSiteConfig_("rakuten").startRow },
    { name: "オークファン", startRow: getSiteConfig_("aucfan").startRow },
    { name: "ヤフオク", startRow: getSiteWriteRange_("yahoo").startRow },
  ];

  sections.forEach(function (section) {
    handleClothingSizeFieldForSection_(sheet, section.startRow, section.name);
  });
}

/**
 * 特定セクションの衣類サイズフィールドを処理
 */
function handleClothingSizeFieldForSection_(sheet, startRow, sectionName) {
  try {
    var clothingSizeRow = startRow + 8; // 服サイズ行
    var range = sheet.getRange(clothingSizeRow, COL_C, 1, MAX_OUTPUT_ITEMS);

    // 現在の値を取得
    var values = range.getValues()[0];

    // 半角英数字に変換
    var convertedValues = values.map(function (value) {
      if (value && typeof value === "string") {
        return toHalfWidthAlphaNum_(value);
      }
      return value;
    });

    // 値を設定
    range.setValues([convertedValues]);

    console.log(sectionName + "の衣類サイズフィールドを処理しました");
  } catch (e) {
    console.error(sectionName + "の衣類サイズ処理エラー:", e.message);
  }
}

/**
 * 全角英数字を半角に変換
 */
function toHalfWidthAlphaNum_(s) {
  if (!s) return "";
  return s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (char) {
    return String.fromCharCode(char.charCodeAt(0) - 0xfee0);
  });
}

// ===== データ取得関数 =====

/**
 * 各サイトのデータを取得する汎用関数
 */
function getSiteData_(sheet, siteConfig) {
  try {
    console.log("URL取得位置:", {
      siteName: siteConfig.siteName,
      row: siteConfig.urlRow,
      col: siteConfig.urlCol,
      cellAddress:
        String.fromCharCode(64 + siteConfig.urlCol) + siteConfig.urlRow,
    });

    var url = sheet.getRange(siteConfig.urlRow, siteConfig.urlCol).getValue();
    console.log("取得したURL値:", url);

    if (!url || url.toString().trim() === "") {
      console.log(siteConfig.siteName + "のURLが設定されていません");
      return [];
    }

    var urlStr = url.toString().trim();

    // URLの妥当性チェック
    if (!urlStr.startsWith("http://") && !urlStr.startsWith("https://")) {
      console.error(siteConfig.siteName + "の無効なURL: " + urlStr);
      return [];
    }

    console.log(siteConfig.siteName + "データ取得開始: " + urlStr);

    var html = siteConfig.fetchFunction(urlStr);
    var items = siteConfig.parseFunction(html);

    // 楽天の場合は売り切れ情報をデバッグ
    if (siteConfig.siteName === "楽天") {
      console.log("=== getSiteData_楽天データ確認 ===");
      console.log("取得した商品数:", items.length);
      items.slice(0, 3).forEach(function (item, index) {
        console.log("getSiteData_商品" + (index + 1) + ":", {
          title: item.title ? item.title.substring(0, 30) + "..." : "なし",
          soldOut: item.soldOut,
          soldOutプロパティ存在: item.hasOwnProperty("soldOut"),
          全プロパティ: Object.keys(item),
        });
      });
    }

    console.log(siteConfig.siteName + "データ取得完了: " + items.length + "件");

    return items;
  } catch (e) {
    console.error(siteConfig.siteName + "データ取得エラー:", e.message);
    return [];
  }
}

// ===== オークファン専用データライター関数 =====

/**
 * オークファン専用のデータライター（別シート対応）
 * 基準行: 3行目（詳細URL）、HTML貼り付け行: 15行目
 */
function writeAucfanDataToSeparateSheet_(sheet, items) {
  console.log(
    "オークファン別シートにデータ配置開始: " + (items ? items.length : 0) + "件"
  );

  // デバッグ情報を追加
  console.log("シート名:", sheet ? sheet.getName() : "なし");
  console.log(
    "MAX_OUTPUT_ITEMS:",
    typeof MAX_OUTPUT_ITEMS !== "undefined" ? MAX_OUTPUT_ITEMS : "未定義"
  );
  console.log("COL_C:", typeof COL_C !== "undefined" ? COL_C : "未定義");

  // 定数が未定義の場合のフォールバック
  if (typeof MAX_OUTPUT_ITEMS === "undefined") {
    var MAX_OUTPUT_ITEMS = 20;
    console.warn("MAX_OUTPUT_ITEMSが未定義のため、デフォルト値20を使用");
  }
  if (typeof COL_C === "undefined") {
    var COL_C = 3;
    console.warn("COL_Cが未定義のため、デフォルト値3を使用");
  }
  if (typeof APPLY_CLIP_WRAP === "undefined") {
    var APPLY_CLIP_WRAP = true;
    console.warn("APPLY_CLIP_WRAPが未定義のため、デフォルト値trueを使用");
  }
  if (typeof MAX_OUTPUT_COLS === "undefined") {
    var MAX_OUTPUT_COLS = 20;
    console.warn("MAX_OUTPUT_COLSが未定義のため、デフォルト値20を使用");
  }

  if (!items || items.length === 0) {
    console.log("オークファンのデータがありません");
    return;
  }

  // 最初の商品データをデバッグ出力
  if (items.length > 0) {
    console.log("最初の商品データ:", {
      title: items[0].title || "なし",
      detailUrl: items[0].detailUrl || "なし",
      price: items[0].price || "なし",
      date: items[0].date || "なし",
    });
  }

  // オークファン専用の行マッピング（基準行3行目）
  var rowMapping = {
    詳細URL: 3, // 詳細URL: 3行目（基準行）
    画像URL: 4, // 画像URL: 4行目
    画像: 5, // 画像: 5行目
    計算に入れるか: 6, // 計算に入れるか: 6行目
    商品名: 7, // 商品名: 7行目
    日付: 8, // 日付: 8行目
    ランク: 9, // ランク: 9行目
    服サイズ: 10, // 服サイズ: 10行目
    カラー: 11, // カラー: 11行目
    価格: 12, // 価格: 12行目
  };

  // 出力件数制限
  var limitedItems = items.slice(0, MAX_OUTPUT_ITEMS);

  // データ配列を準備
  var rowArrays = {
    詳細URL: limitedItems.map((it) => it.detailUrl || ""),
    画像URL: limitedItems.map((it) => it.imageUrl || ""),
    画像: limitedItems.map(() => ""), // 画像行は空のまま（画像表示用の数式が入る）
    計算に入れるか: limitedItems.map(() => ""), // 空のまま（手動入力用）
    商品名: limitedItems.map((it) => it.title || ""),
    日付: limitedItems.map((it) => it.date || ""),
    ランク: limitedItems.map((it) => normalizeRankForWrite_(it.rank || "")),
    服サイズ: limitedItems.map(() => ""), // 空のまま（手動入力用）
    カラー: limitedItems.map(() => ""), // 空のまま（手動入力用）
    引用サイト: limitedItems.map((it) => it.source || ""),
    価格: limitedItems.map((it) => it.price || ""),
    ショップ: limitedItems.map((it) => it.shop || ""),
  };

  // 各ラベルに対応する行にデータを配置
  Object.keys(rowMapping).forEach(function (label) {
    var targetRow = rowMapping[label];
    var data = rowArrays[label];

    console.log(
      `処理中: ${label} (行${targetRow}) - データ数: ${data ? data.length : 0}`
    );

    // 画像行はスキップ（画像表示用の数式が入るため、データは書き込まない）
    if (label === "画像") {
      console.log(
        "オークファンの「" + label + "」行はスキップしました（数式保持のため）"
      );
      return;
    }

    if (data && data.length > 0) {
      try {
        // データ配列を準備
        var values = [];
        for (var i = 0; i < MAX_OUTPUT_ITEMS; i++) {
          var value = data && i < data.length ? data[i] : "";

          // ランクの場合、有効な値のみ設定
          if (label === "ランク") {
            var validRanks = ["N", "S", "SA", "A", "AB", "B", "BC", "C"];
            if (value && validRanks.indexOf(value.toUpperCase()) > -1) {
              values.push(value.toUpperCase());
            } else {
              values.push("");
            }
          } else {
            values.push(value);
          }
        }

        var range = sheet.getRange(targetRow, COL_C, 1, MAX_OUTPUT_ITEMS);
        console.log(
          `書き込み範囲: 行${targetRow}, 列${COL_C}, 幅${MAX_OUTPUT_ITEMS}`
        );
        console.log(`書き込みデータ例:`, values.slice(0, 3));

        // ランクの場合は特別な処理
        if (label === "ランク") {
          console.log(
            "オークファンのランクデータ:",
            values.filter((v) => v !== "").join(", ")
          );

          // ランクの場合は個別にセルに値を設定
          for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
            var value = values[col];
            var cellRange = sheet.getRange(targetRow, COL_C + col);
            cellRange.setValue(value);
            if (value && value !== "") {
              console.log(
                `ランク書き込み: 行${targetRow}, 列${COL_C + col}, 値: ${value}`
              );
            }
          }
        } else if (label === "画像URL") {
          // 画像URL行の場合は数式を保持してデータを設定
          console.log(
            "オークファンの画像URLデータ:",
            values.filter((v) => v !== "").slice(0, 3)
          );

          var range = sheet.getRange(targetRow, COL_C, 1, MAX_OUTPUT_ITEMS);
          var formulas = range.getFormulas()[0];
          var newValues = [[]];

          for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
            var formula = formulas[col];
            var value = values[col];

            if (formula) {
              newValues[0][col] = formula;
            } else if (value && value !== "") {
              newValues[0][col] = value;
            } else {
              newValues[0][col] = "";
            }
          }

          range.setValues(newValues);
        } else if (label === "服サイズ" || label === "カラー") {
          // 服サイズとカラーの場合は背景色も保持
          console.log(
            "オークファンの" + label + "データ（背景色保持）:",
            values.filter((v) => v !== "").slice(0, 3)
          );

          for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
            var cellRange = sheet.getRange(targetRow, COL_C + col);
            var validation = cellRange.getDataValidation();
            var bgColor = cellRange.getBackground();
            var value = values[col];

            if (value !== null && value !== undefined) {
              cellRange.setValue(value);
            }

            if (validation) {
              cellRange.setDataValidation(validation);
            }

            if (bgColor) {
              cellRange.setBackground(bgColor);
            }
          }
        } else if (label === "計算に入れるか") {
          // チェックボックス系は一括書き込み（データ入力規則を保持）
          console.log(
            "オークファンの" + label + "データ（チェックボックス）:",
            values.slice(0, 5)
          );
          range.setValues([values]);
        } else {
          // その他の場合は従来の処理
          console.log(`${label}の通常書き込み処理開始`);
          var validationRules = [];
          var bgColors = [];

          for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
            var cellRange = sheet.getRange(targetRow, COL_C + col);
            var validation = cellRange.getDataValidation();
            var bgColor = cellRange.getBackground();
            validationRules.push(validation);
            bgColors.push(bgColor);
            if (validation) {
              cellRange.clearDataValidations();
            }
          }

          console.log(`${label}の値を一括書き込み中...`);
          range.setValues([values]);
          console.log(`${label}の書き込み完了`);

          for (var col = 0; col < MAX_OUTPUT_ITEMS; col++) {
            var cellRange = sheet.getRange(targetRow, COL_C + col);
            if (validationRules[col]) {
              cellRange.setDataValidation(validationRules[col]);
            }
            if (bgColors[col]) {
              cellRange.setBackground(bgColors[col]);
            }
          }
        }

        // 折返し抑止
        if (APPLY_CLIP_WRAP) {
          var outLen = Math.min(limitedItems.length, MAX_OUTPUT_COLS);
          if (outLen > 0) {
            sheet
              .getRange(targetRow, COL_C, 1, outLen)
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          }
        }
      } catch (e) {
        console.error(
          "オークファンの「" + label + "」書き込みエラー:",
          e.message
        );
        console.error("エラースタック:", e.stack);
        console.error("エラー発生時のデータ:", values.slice(0, 3));
      }
    } else if (label === "ランク") {
      console.log("オークファンのランクデータは空のためスキップしました");
    } else {
      console.log(
        `${label}のデータが空またはなし - データ数: ${data ? data.length : 0}`
      );
    }
  });

  console.log(
    "オークファン別シートにデータ配置完了: " + limitedItems.length + "件"
  );
}

/**
 * オークファン専用のセクションクリア（別シート対応）
 */
function clearAucfanSeparateSheet_(sheet) {
  try {
    console.log("オークファン別シートのクリア開始");

    // オークファン専用の行マッピング（基準行3行目）
    var rowMapping = {
      詳細URL: 3, // 詳細URL: 3行目（基準行）
      画像URL: 4, // 画像URL: 4行目
      画像: 5, // 画像: 5行目（数式保持）
      計算に入れるか: 6, // 計算に入れるか: 6行目
      商品名: 7, // 商品名: 7行目
      日付: 8, // 日付: 8行目
      ランク: 9, // ランク: 9行目
      服サイズ: 10, // 服サイズ: 10行目
      カラー: 11, // カラー: 11行目
      価格: 12, // 価格: 12行目
    };

    // 一括クリア処理
    clearSectionBatch_(sheet, rowMapping);

    console.log("オークファン別シートのクリア完了");
  } catch (e) {
    console.error("オークファン別シートのクリアエラー:", e.message);
  }
}

/**
 * オークファン専用のHTMLからデータ取得（別シート対応）
 * HTML貼り付け行: 15行目
 */
function getAucfanDataFromSeparateSheet_(sheet) {
  try {
    // 15行目以降にHTMLがあるかチェック
    var htmlFromSheet = readHtmlFromRow_(sheet, 15);

    if (htmlFromSheet) {
      // HTMLソースがオークファンかどうかを判定
      var source = detectSource_(htmlFromSheet);
      if (source === "オークファン" || source === "aucfan") {
        console.log("15行目以降のHTMLからオークファンデータを取得します");
        var items = parseAucfanFromHtml_(htmlFromSheet);
        console.log("オークファンデータを取得しました:", items.length + "件");
        return items;
      } else {
        console.log("15行目以降のHTMLはオークファンではありません:", source);
      }
    }

    console.log("オークファンのHTMLが見つかりません（15行目以降を確認）");
    return [];
  } catch (e) {
    console.warn("オークファン別シートからのデータ取得に失敗:", e.message);
    return [];
  }
}

// ===== 個別サイト用ラッパー関数 =====

/**
 * SBAデータを取得
 */
function getSbaData_(sheet) {
  try {
    // 設定値キャッシュをクリアして最新の設定を取得
    clearConfigCache_();
    var siteConfig = getSiteConfig_("sba");
    var urlPosition = getSiteUrlPosition_("sba");

    console.log("SBA設定値:", {
      name: siteConfig.name,
      urlRow: urlPosition.row,
      urlCol: urlPosition.col,
    });

    return getSiteData_(sheet, {
      siteName: siteConfig.name,
      urlRow: urlPosition.row,
      urlCol: urlPosition.col,
      fetchFunction: fetchStarBuyersHtml_,
      parseFunction: parseStarBuyersFromHtml_,
    });
  } catch (e) {
    console.warn("SBA設定値読み取りエラー、フォールバック値を使用:", e.message);
    return getSiteData_(sheet, {
      siteName: "SBA",
      urlRow: 4,
      urlCol: 2,
      fetchFunction: fetchStarBuyersHtml_,
      parseFunction: parseStarBuyersFromHtml_,
    });
  }
}

/**
 * 楽天データを取得
 */
function getRakutenData_(sheet) {
  try {
    var siteConfig = getSiteConfig_("rakuten");
    var urlPosition = getSiteUrlPosition_("rakuten");

    // URL取得
    var url = sheet.getRange(urlPosition.row, urlPosition.col).getValue();
    if (!url || String(url).trim() === "") {
      console.log("楽天のURLが設定されていません");
      return [];
    }

    var urlStr = String(url).trim();
    if (!urlStr.startsWith("http://") && !urlStr.startsWith("https://")) {
      console.error("楽天の無効なURL: " + urlStr);
      return [];
    }

    console.log("楽天（ページネーション対応）データ取得開始: " + urlStr);
    // ページネーション + 売り切れフィルタリング
    var soldOutItems = scrapeRakutenWithPagination_(urlStr);
    console.log("楽天（ページネーション対応）取得完了: " + soldOutItems.length + "件");
    return soldOutItems;
  } catch (e) {
    console.warn(
      "楽天設定値読み取りエラー、フォールバック値を使用:",
      e.message
    );
    // フォールバック: B30からURLを取得してページネーション処理
    try {
      var fallbackUrl = sheet.getRange(30, 2).getValue();
      var urlStr = String(fallbackUrl || "").trim();
      if (!urlStr || (!urlStr.startsWith("http://") && !urlStr.startsWith("https://"))) {
        console.error("楽天フォールバックURLが無効です");
        return [];
      }
      console.log("楽天（フォールバック・ページネーション）開始: " + urlStr);
      return scrapeRakutenWithPagination_(urlStr);
    } catch (ee) {
      console.error("楽天フォールバック処理エラー:", ee.message);
      return [];
    }
  }
}

/**
 * オークファンデータを取得
 */
function getAucfanData_(sheet) {
  try {
    var siteConfig = getSiteConfig_("aucfan");
    var urlPosition = getSiteUrlPosition_("aucfan");

    // オークファン専用の取得関数を使用（HTMLまたはURLから）
    return getAucfanDataFromSheetOrUrl_(
      sheet,
      urlPosition.row,
      urlPosition.col
    );
  } catch (e) {
    console.warn(
      "オークファン設定値読み取りエラー、フォールバック値を使用:",
      e.message
    );
    // フォールバック値でも専用関数を使用
    return getAucfanDataFromSheetOrUrl_(sheet, 110, 2);
  }
}

/**
 * エコリングデータを取得
 */
function getEcoringData_(sheet) {
  try {
    var siteConfig = getSiteConfig_("ecoring");
    var urlPosition = getSiteUrlPosition_("ecoring");

    return getSiteData_(sheet, {
      siteName: siteConfig.name,
      urlRow: urlPosition.row,
      urlCol: urlPosition.col,
      fetchFunction: fetchEcoringHtml_,
      parseFunction: parseEcoringFromHtml_,
    });
  } catch (e) {
    console.warn(
      "エコリング設定値読み取りエラー、フォールバック値を使用:",
      e.message
    );
    return getSiteData_(sheet, {
      siteName: "エコリング",
      urlRow: 17,
      urlCol: 2,
      fetchFunction: fetchEcoringHtml_,
      parseFunction: parseEcoringFromHtml_,
    });
  }
}

/**
 * ヤフオクデータを取得
 */
function getYahooAuctionData_(sheet) {
  try {
    var siteConfig = getSiteConfig_("yahoo");
    var urlPosition = getSiteUrlPosition_("yahoo");

    return getSiteData_(sheet, {
      siteName: siteConfig.name,
      urlRow: urlPosition.row,
      urlCol: urlPosition.col,
      fetchFunction: fetchYahooAuctionHtml_,
      parseFunction: parseYahooAuctionFromHtml_,
    });
  } catch (e) {
    console.warn(
      "ヤフオク設定値読み取りエラー、フォールバック値を使用:",
      e.message
    );
    return getSiteData_(sheet, {
      siteName: "ヤフオク",
      urlRow: 44,
      urlCol: 2,
      fetchFunction: fetchYahooAuctionHtml_,
      parseFunction: parseYahooAuctionFromHtml_,
    });
  }
}
