/**
 * PJ PL管理 — 書式設定
 * create_pj_pl_v2.py の487-623行目を移植
 */

// ─── ダッシュボード書式 ───
function formatDashboard(sheet, pjCount) {
  var lastCol = 11;

  // タイトル行
  var titleRange = sheet.getRange(1, 1, 1, lastCol);
  titleRange.setBackground(CLR_DARK_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold").setFontSize(14);

  // ヘッダー行
  var headerRange = sheet.getRange(3, 1, 1, lastCol);
  headerRange.setBackground(CLR_MID_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold");

  // データ行の金額書式
  sheet.getRange(4, 5, pjCount + 2, 5).setNumberFormat("¥#,##0"); // E-I列
  sheet.getRange(4, 10, pjCount + 2, 1).setNumberFormat("0.0%");  // J列（粗利率）

  // 交互の背景色
  for (var i = 0; i < pjCount; i++) {
    var bg = (i % 2 === 0) ? CLR_LIGHT_GRAY : CLR_WHITE;
    sheet.getRange(4 + i, 1, 1, lastCol).setBackground(bg);
  }

  // 合計行
  sheet.getRange(4 + pjCount + 1, 1, 1, lastCol).setBackground(CLR_LIGHT_YELLOW).setFontWeight("bold");

  // 列幅
  var widths = [90, 150, 230, 100, 120, 110, 110, 110, 110, 80, 100];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  // 行固定
  sheet.setFrozenRows(3);
}

// ─── 個別PLシート書式 ───
function formatPlSheet(sheet, N) {
  var totalCol = C_START + N + 1; // PL合計列（1-indexed）
  var laborTotalCol = L_START + N + 1; // 工数合計列（1-indexed）
  var maxCol = Math.max(totalCol, laborTotalCol);

  // タイトル行
  sheet.getRange(R(R_TITLE) + 0, 1, 1, maxCol).setBackground(CLR_DARK_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold").setFontSize(12);

  // 基本情報ラベル
  sheet.getRange(R(R_INFO), 1, 3, 1).setFontWeight("bold");
  // 基本情報の金額
  sheet.getRange(R(R_INFO + 1), 2, 1, 1).setNumberFormat("¥#,##0");
  sheet.getRange(R(R_INFO + 1), 5, 1, 1).setNumberFormat("¥#,##0");

  // サマリー
  sheet.getRange(R(R_SUMMARY_H), 1, 1, 6).setBackground(CLR_LIGHT_GREEN).setFontWeight("bold");
  sheet.getRange(R(R_SUMMARY_V), 2, 1, 3).setNumberFormat("¥#,##0");
  sheet.getRange(R(R_SUMMARY_V), 5, 1, 1).setNumberFormat("0.0%");
  sheet.getRange(R(R_SUMMARY_V), 6, 1, 1).setNumberFormat("¥#,##0");

  // 月別ヘッダー
  sheet.getRange(R(R_MONTHLY_H), 1, 1, totalCol).setBackground(CLR_LIGHT_BLUE).setFontWeight("bold");

  // 月別の金額（売上〜外注費）
  sheet.getRange(R(R_SALES), C_START + 1, R_EXTERNAL - R_SALES + 1, N + 1).setNumberFormat("¥#,##0");
  // 粗利
  sheet.getRange(R(R_GROSS), C_START + 1, 1, N + 1).setNumberFormat("¥#,##0");
  sheet.getRange(R(R_GROSS), 1, 1, totalCol).setBackground(CLR_LIGHT_YELLOW).setFontWeight("bold");
  // 粗利率
  sheet.getRange(R(R_GROSS_PCT), C_START + 1, 1, N + 1).setNumberFormat("0.0%");

  // 工数ヘッダー
  sheet.getRange(R(R_LABOR_H), 1, 1, laborTotalCol).setBackground(CLR_LIGHT_BLUE).setFontWeight("bold");
  // 工数の配分率（%表示）— D列以降
  sheet.getRange(R(R_LABOR_START), L_START + 1, MAX_MEMBERS, N).setNumberFormat("0.0%");
  // 工数合計金額列
  sheet.getRange(R(R_LABOR_START), laborTotalCol, MAX_MEMBERS + 1, 1).setNumberFormat("¥#,##0");
  // 工数小計行
  sheet.getRange(R(R_LABOR_SUB), 1, 1, laborTotalCol).setBackground(CLR_LIGHT_GRAY).setFontWeight("bold").setNumberFormat("¥#,##0");
  // 単価列
  sheet.getRange(R(R_LABOR_START), 3, MAX_MEMBERS, 1).setNumberFormat("¥#,##0");

  // 外注ヘッダー
  sheet.getRange(R(R_VENDOR_H), 1, 1, totalCol).setBackground(CLR_LIGHT_BLUE).setFontWeight("bold");
  // 外注金額
  sheet.getRange(R(R_VENDOR_START), C_START + 1, MAX_VENDORS + 1, N + 1).setNumberFormat("¥#,##0");
  // 外注小計行
  sheet.getRange(R(R_VENDOR_SUB), 1, 1, totalCol).setBackground(CLR_LIGHT_GRAY).setFontWeight("bold");

  // 列幅
  sheet.setColumnWidth(1, 100); // A列
  sheet.setColumnWidth(2, 120); // B列
  for (var c = 3; c <= maxCol; c++) {
    sheet.setColumnWidth(c, 110);
  }

  // 行固定
  sheet.setFrozenRows(R(R_MONTHLY_H));
  sheet.setFrozenColumns(2);
}

// ─── 工数管理シート書式 ───
function formatLaborMgmt(sheet) {
  // タイトル
  sheet.getRange(1, 1, 1, 3).setBackground(CLR_DARK_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold").setFontSize(14);
  // ヘッダー
  sheet.getRange(3, 1, 1, 3).setBackground(CLR_MID_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold");
  // 金額列
  sheet.getRange(4, 3, 200, 1).setNumberFormat("¥#,##0");
  // 列幅
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 120);
  // 行固定
  sheet.setFrozenRows(3);
}

// ─── 外注費管理シート書式 ───
function formatVendorMgmt(sheet) {
  sheet.getRange(1, 1, 1, 3).setBackground(CLR_DARK_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold").setFontSize(14);
  sheet.getRange(3, 1, 1, 3).setBackground(CLR_MID_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold");
  sheet.getRange(4, 3, 200, 1).setNumberFormat("¥#,##0");
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 120);
  sheet.setFrozenRows(3);
}

// ─── PJ台帳書式 ───
function formatLedger(sheet) {
  sheet.getRange(1, 1, 1, 9).setBackground(CLR_DARK_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold").setFontSize(14);
  sheet.getRange(3, 1, 1, 9).setBackground(CLR_MID_BLUE).setFontColor(CLR_WHITE).setFontWeight("bold");
  sheet.getRange(4, 5, 30, 1).setNumberFormat("¥#,##0"); // 契約額
  sheet.getRange(4, 9, 30, 1).setNumberFormat("¥#,##0"); // 月額売上
  var widths = [90, 150, 230, 100, 120, 110, 110, 80, 110];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  sheet.setFrozenRows(3);
}
