/**
 * 次の値が今の値より大きい回数（=降順崩れ）を数える。
 * 0や非数は無視。通貨記号・カンマ・全角数字にも対応。
 * @customfunction
 */
function COUNT_DESC_BREAKS(range) {
    const arr1d = _flatten1D_(range);
    const nums = _toNums_(arr1d).filter(v => v > 0);
    return _countAscents_(nums);
  }
  
  /**
   * 2D範囲の各行ごとに降順崩れ回数を返す（縦ベクトル）。
   * @customfunction
   */
  function COUNT_DESC_BREAKS_BY_ROW(range) {
    const out = range.map(row => {
      const nums = _toNums_(row).filter(v => v > 0);
      return [_countAscents_(nums)];
    });
    return out;
  }
  
  /* ==== ヘルパー ==== */
  
  // 2D配列を1Dに
  function _flatten1D_(a2d) {
    return a2d.reduce((acc, row) => acc.concat(row), []);
  }
  
  // 全角→半角、通貨・カンマ除去して数値化
  function _toNums_(arr) {
    return arr.map(v => {
      if (typeof v === 'number') return v;
      if (v === null || v === '') return NaN;
      let s = String(v);
      // 全角英数→半角
      s = s.replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
      // 数字・符号以外除去（通貨記号・カンマ等を落とす）
      s = s.replace(/[^\d\.\-]/g, '');
      const n = parseFloat(s);
      return Number.isFinite(n) ? n : NaN;
    }).filter(n => Number.isFinite(n));
  }
  
  // 隣接ペアで「次>今」をカウント
  function _countAscents_(nums) {
    let c = 0;
    for (let i = 0; i < nums.length - 1; i++) {
      if (nums[i + 1] > nums[i]) c++;
    }
    return c;
  }
  