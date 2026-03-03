/**
 * PJ PL管理 — ダッシュボード・集約シート構築
 * create_pj_pl_v2.py の351-470行目を移植
 */

// ─── ダッシュボード ───
function buildDashboard(sheet, pjList, plSheets) {
  var rows = [];

  // タイトル
  rows.push(["PJ PL管理 ダッシュボード", "", "", "", "", "", "", "", "", "", ""]);
  rows.push(new Array(11).fill(""));

  // ヘッダー
  rows.push(["PJコード", "PJ名", "会社名", "商材", "契約額", "月額売上", "売上合計", "原価合計", "粗利", "粗利率", "ステータス"]);

  // 各PJの行
  pjList.forEach(function(pj) {
    var title = plSheetTitle(pj.name);
    rows.push([
      pj.code,
      pj.name,
      pj.company,
      pj.type,
      pj.amount,
      pj.monthly,
      "='" + title + "'!B" + R(R_SUMMARY_V),
      "='" + title + "'!C" + R(R_SUMMARY_V),
      "='" + title + "'!D" + R(R_SUMMARY_V),
      "='" + title + "'!E" + R(R_SUMMARY_V),
      "",
    ]);
  });

  // 空行
  rows.push(new Array(11).fill(""));

  // 合計行
  var startR = 4; // データ開始行（1-indexed）
  var endR = startR + pjList.length - 1;
  var totalR = endR + 2;
  var totalRow = [
    "合計", "", "", "",
    "=SUM(E" + startR + ":E" + endR + ")",
    "=SUM(F" + startR + ":F" + endR + ")",
    "=SUM(G" + startR + ":G" + endR + ")",
    "=SUM(H" + startR + ":H" + endR + ")",
    "=SUM(I" + startR + ":I" + endR + ")",
    '=IF(G' + totalR + '=0,"-",I' + totalR + '/G' + totalR + ')',
    "",
  ];
  rows.push(totalRow);

  // 列数をそろえる
  var maxCols = 11;
  rows.forEach(function(row) {
    while (row.length < maxCols) row.push("");
  });

  sheet.getRange(1, 1, rows.length, maxCols).setValues(rows);
}

// ─── 工数管理 ───
function buildLaborMgmt(sheet, pjList, plSheets) {
  var rows = [];

  rows.push(["工数管理（自動集約）", "", "単価: ¥800,000"]);
  rows.push(["", "", ""]);
  rows.push(["PJ名", "メンバー", "合計工数(円)"]);

  pjList.forEach(function(pj) {
    var title = plSheetTitle(pj.name);
    var pl = plSheets[pj.key];
    var N = pl.monthLabels.length;

    for (var mIdx = 0; mIdx < MAX_MEMBERS; mIdx++) {
      var r = R(R_LABOR_START + mIdx);
      rows.push([
        pj.name,
        "='" + title + "'!B" + r,
        "='" + title + "'!" + colLetter(L_START + N) + r,
      ]);
    }
    // 小計行
    var subR = R(R_LABOR_SUB);
    rows.push([
      "  " + pj.name + " 小計", "",
      "='" + title + "'!" + colLetter(L_START + N) + subR,
    ]);
    rows.push(["", "", ""]);
  });

  // 列数をそろえる
  rows.forEach(function(row) {
    while (row.length < 3) row.push("");
  });

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}

// ─── 外注費管理 ───
function buildVendorMgmt(sheet, pjList, plSheets) {
  var rows = [];

  rows.push(["外注費管理（自動集約）", "", ""]);
  rows.push(["", "", ""]);
  rows.push(["PJ名", "ベンダー", "合計"]);

  pjList.forEach(function(pj) {
    var title = plSheetTitle(pj.name);
    var pl = plSheets[pj.key];
    var N = pl.monthLabels.length;
    var cnVendor = colLetter(C_START + N);

    for (var vIdx = 0; vIdx < MAX_VENDORS; vIdx++) {
      var r = R(R_VENDOR_START + vIdx);
      rows.push([
        pj.name,
        "='" + title + "'!B" + r,
        "='" + title + "'!" + cnVendor + r,
      ]);
    }
    // 小計行
    var subR = R(R_VENDOR_SUB);
    rows.push([
      "  " + pj.name + " 小計", "",
      "='" + title + "'!" + cnVendor + subR,
    ]);
    rows.push(["", "", ""]);
  });

  rows.forEach(function(row) {
    while (row.length < 3) row.push("");
  });

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}

// ─── PJ台帳 ───
function buildLedger(sheet, pjList) {
  var rows = [];

  rows.push(["PJ台帳", "", "", "", "", "", "", "", ""]);
  rows.push(new Array(9).fill(""));
  rows.push(["PJコード", "PJ名", "会社名", "商材", "契約額", "契約開始", "契約終了", "期間", "月額売上"]);

  pjList.forEach(function(pj) {
    rows.push([
      pj.code, pj.name, pj.company, pj.type,
      pj.amount, pj.start, pj.end,
      pj.months ? pj.months + "ヶ月" : "", pj.monthly,
    ]);
  });

  rows.forEach(function(row) {
    while (row.length < 9) row.push("");
  });

  sheet.getRange(1, 1, rows.length, 9).setValues(rows);
}
