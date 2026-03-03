/**
 * PJ PL管理 — 個別PLシート構築
 * create_pj_pl_v2.py の182-345行目を移植
 */

// 1つのPLシートを構築する
function buildPlSheet(sheet, pj, plData) {
  var months = plData.monthLabels;
  var N = months.length;
  var C_END = C_START + N;
  var cSum = colLetter(C_END); // 合計列のアルファベット
  var c0 = colLetter(C_START);
  var cn = colLetter(C_END - 1);

  // 全行を組み立て（R_VENDOR_SUB + 2行分）
  var totalRows = R_VENDOR_SUB + 2;
  var totalCols = C_END + 1; // 合計列まで
  // 工数セクションは列が1つ多い（単価列）ので最大列数を計算
  var laborCols = L_START + N + 1; // 合計(円)列まで
  var maxCols = Math.max(totalCols, laborCols);

  var rows = [];
  for (var i = 0; i < totalRows; i++) {
    rows.push(new Array(maxCols).fill(""));
  }

  // ─── Row 0: タイトル ───
  rows[R_TITLE][0] = "PL: " + pj.name;

  // ─── Row 1-3: 基本情報 ───
  rows[R_INFO][0] = "会社名"; rows[R_INFO][1] = pj.company;
  rows[R_INFO][3] = "商材"; rows[R_INFO][4] = pj.type;
  rows[R_INFO + 1][0] = "契約額"; rows[R_INFO + 1][1] = pj.amount;
  rows[R_INFO + 1][3] = "月額売上"; rows[R_INFO + 1][4] = pj.monthly;
  var period = pj.start ? (pj.start + "～" + pj.end) : "";
  rows[R_INFO + 2][0] = "契約期間"; rows[R_INFO + 2][1] = period;
  rows[R_INFO + 2][3] = "チーム"; rows[R_INFO + 2][4] = plData.team || "";

  // ─── Row 5: サマリーヘッダー ───
  rows[R_SUMMARY_H][0] = "サマリー";
  rows[R_SUMMARY_H][1] = "売上高";
  rows[R_SUMMARY_H][2] = "原価";
  rows[R_SUMMARY_H][3] = "粗利";
  rows[R_SUMMARY_H][4] = "粗利率";
  rows[R_SUMMARY_H][5] = "営利";

  // ─── Row 6: サマリー値（数式） ───
  var rS = R(R_SALES), rC = R(R_COST), rG = R(R_GROSS);
  rows[R_SUMMARY_V][1] = "=SUM(" + c0 + rS + ":" + cn + rS + ")";
  rows[R_SUMMARY_V][2] = "=SUM(" + c0 + rC + ":" + cn + rC + ")";
  rows[R_SUMMARY_V][3] = "=B" + R(R_SUMMARY_V) + "-C" + R(R_SUMMARY_V);
  rows[R_SUMMARY_V][4] = '=IF(B' + R(R_SUMMARY_V) + '=0,"-",D' + R(R_SUMMARY_V) + '/B' + R(R_SUMMARY_V) + ')';
  rows[R_SUMMARY_V][5] = "=D" + R(R_SUMMARY_V);

  // ─── Row 8: 月別ヘッダー ───
  rows[R_MONTHLY_H][0] = "月別計上";
  rows[R_MONTHLY_H][1] = "科目";
  for (var i = 0; i < N; i++) rows[R_MONTHLY_H][C_START + i] = months[i];
  rows[R_MONTHLY_H][C_END] = "合計";

  // ─── Row 9: 売上 ───
  rows[R_SALES][1] = "売上高";
  for (var i = 0; i < N; i++) rows[R_SALES][C_START + i] = plData.sales[i];
  rows[R_SALES][C_END] = "=SUM(" + c0 + R(R_SALES) + ":" + cn + R(R_SALES) + ")";

  // ─── Row 10: 原価（= 社内工数 + 外注費） ───
  rows[R_COST][1] = "原価";
  for (var i = 0; i < N; i++) {
    var c = colLetter(C_START + i);
    rows[R_COST][C_START + i] = "=" + c + R(R_INTERNAL) + "+" + c + R(R_EXTERNAL);
  }
  rows[R_COST][C_END] = "=SUM(" + c0 + R(R_COST) + ":" + cn + R(R_COST) + ")";

  // ─── Row 11: 社内工数（= 工数小計を参照） ───
  rows[R_INTERNAL][1] = "  社内工数";
  for (var i = 0; i < N; i++) {
    var lc = colLetter(L_START + i); // 工数小計のD列以降を参照
    rows[R_INTERNAL][C_START + i] = "=" + lc + R(R_LABOR_SUB);
  }
  rows[R_INTERNAL][C_END] = "=SUM(" + c0 + R(R_INTERNAL) + ":" + cn + R(R_INTERNAL) + ")";

  // ─── Row 12: 外注費（= 外注小計を参照） ───
  rows[R_EXTERNAL][1] = "  外注費";
  for (var i = 0; i < N; i++) {
    var c = colLetter(C_START + i);
    rows[R_EXTERNAL][C_START + i] = "=" + c + R(R_VENDOR_SUB);
  }
  rows[R_EXTERNAL][C_END] = "=SUM(" + c0 + R(R_EXTERNAL) + ":" + cn + R(R_EXTERNAL) + ")";

  // ─── Row 13: 粗利 ───
  rows[R_GROSS][1] = "粗利";
  for (var i = 0; i < N; i++) {
    var c = colLetter(C_START + i);
    rows[R_GROSS][C_START + i] = "=" + c + R(R_SALES) + "-" + c + R(R_COST);
  }
  rows[R_GROSS][C_END] = "=SUM(" + c0 + R(R_GROSS) + ":" + cn + R(R_GROSS) + ")";

  // ─── Row 14: 粗利率 ───
  rows[R_GROSS_PCT][1] = "粗利率";
  for (var i = 0; i < N; i++) {
    var c = colLetter(C_START + i);
    rows[R_GROSS_PCT][C_START + i] = '=IF(' + c + R(R_SALES) + '=0,"-",' + c + R(R_GROSS) + '/' + c + R(R_SALES) + ')';
  }
  rows[R_GROSS_PCT][C_END] = '=IF(' + cSum + R(R_SALES) + '=0,"-",' + cSum + R(R_GROSS) + '/' + cSum + R(R_SALES) + ')';

  // ─── Row 16: 工数ヘッダー ───
  rows[R_LABOR_H][0] = "社内工数";
  rows[R_LABOR_H][1] = "メンバー";
  rows[R_LABOR_H][2] = "単価";
  for (var i = 0; i < N; i++) rows[R_LABOR_H][L_START + i] = months[i];
  rows[R_LABOR_H][L_START + N] = "合計(円)";

  // ─── Row 17-21: メンバー（最大5名） ───
  var members = plData.members || [];
  for (var mIdx = 0; mIdx < MAX_MEMBERS; mIdx++) {
    var r = R_LABOR_START + mIdx;
    rows[r][2] = UNIT_PRICE; // 単価
    if (mIdx < members.length) {
      var m = members[mIdx];
      rows[r][1] = m.name;
      for (var i = 0; i < N; i++) {
        rows[r][L_START + i] = (i < m.pct.length) ? m.pct[i] / 100 : 0;
      }
      // 合計(円) = SUM(単価 × 各月配分率)
      var sumParts = [];
      for (var i = 0; i < N; i++) {
        var c = colLetter(L_START + i);
        sumParts.push("C" + R(r) + "*" + c + R(r));
      }
      rows[r][L_START + N] = "=" + sumParts.join("+");
    } else {
      // 空行
      for (var i = 0; i < N; i++) rows[r][L_START + i] = 0;
      rows[r][L_START + N] = "=0";
    }
  }

  // ─── Row 22: 工数小計 ───
  rows[R_LABOR_SUB][1] = "小計";
  for (var i = 0; i < N; i++) {
    var c = colLetter(L_START + i);
    var parts = [];
    for (var mIdx = 0; mIdx < MAX_MEMBERS; mIdx++) {
      var mr = R(R_LABOR_START + mIdx);
      parts.push("$C" + mr + "*" + c + mr);
    }
    rows[R_LABOR_SUB][L_START + i] = "=" + parts.join("+");
  }
  var lC0 = colLetter(L_START);
  var lCn = colLetter(L_START + N - 1);
  rows[R_LABOR_SUB][L_START + N] = "=SUM(" + lC0 + R(R_LABOR_SUB) + ":" + lCn + R(R_LABOR_SUB) + ")";

  // ─── Row 24: 外注ヘッダー ───
  rows[R_VENDOR_H][0] = "外注費";
  rows[R_VENDOR_H][1] = "ベンダー";
  for (var i = 0; i < N; i++) rows[R_VENDOR_H][C_START + i] = months[i];
  rows[R_VENDOR_H][C_END] = "合計";

  // ─── Row 25-29: ベンダー（最大5社） ───
  var vendors = plData.vendors || [];
  for (var vIdx = 0; vIdx < MAX_VENDORS; vIdx++) {
    var r = R_VENDOR_START + vIdx;
    if (vIdx < vendors.length) {
      var v = vendors[vIdx];
      rows[r][1] = v.name;
      for (var i = 0; i < N; i++) {
        rows[r][C_START + i] = (i < v.amounts.length) ? v.amounts[i] : 0;
      }
    } else {
      for (var i = 0; i < N; i++) rows[r][C_START + i] = 0;
    }
    rows[r][C_END] = "=SUM(" + c0 + R(r) + ":" + cn + R(r) + ")";
  }

  // ─── Row 30: 外注小計 ───
  rows[R_VENDOR_SUB][1] = "小計";
  for (var i = 0; i < N; i++) {
    var c = colLetter(C_START + i);
    rows[R_VENDOR_SUB][C_START + i] = "=SUM(" + c + R(R_VENDOR_START) + ":" + c + R(R_VENDOR_START + MAX_VENDORS - 1) + ")";
  }
  rows[R_VENDOR_SUB][C_END] = "=SUM(" + c0 + R(R_VENDOR_SUB) + ":" + cn + R(R_VENDOR_SUB) + ")";

  // 一括書き込み
  sheet.getRange(1, 1, rows.length, maxCols).setValues(rows);
}
