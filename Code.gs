/**
 * PJ PL管理 — メイン制御
 * スプレッドシートのメニューから操作する
 */

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("PJ PL管理")
    .addItem("全シート再構築", "rebuildAll")
    .addItem("個別PLのみ再構築", "rebuildPLSheets")
    .addItem("ダッシュボードのみ更新", "rebuildDashboardOnly")
    .addToUi();
}

// ─── 全シート再構築 ───
function rebuildAll() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "全シート再構築",
    "すべてのシートを初期化して再構築します。\n手動で追加したデータは消えます。\n\n続行しますか？",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pjList = getPjList();
  var plSheets = getPlSheetsWithDefaults();

  // 進捗表示
  ss.toast("シートを準備中...", "PJ PL管理", -1);

  // 既存シートを全削除して再作成
  deleteAndRecreateSheets_(ss, pjList);

  // 個別PLシートを構築
  ss.toast("PLシートを構築中...", "PJ PL管理", -1);
  pjList.forEach(function(pj) {
    var title = plSheetTitle(pj.name);
    var sheet = ss.getSheetByName(title);
    if (sheet) {
      buildPlSheet(sheet, pj, plSheets[pj.key]);
      formatPlSheet(sheet, plSheets[pj.key].monthLabels.length);
    }
  });

  // ダッシュボード構築
  ss.toast("ダッシュボードを構築中...", "PJ PL管理", -1);
  var dashSheet = ss.getSheetByName("ダッシュボード");
  if (dashSheet) {
    buildDashboard(dashSheet, pjList, plSheets);
    formatDashboard(dashSheet, pjList.length);
  }

  // 工数管理
  var laborSheet = ss.getSheetByName("工数管理");
  if (laborSheet) {
    buildLaborMgmt(laborSheet, pjList, plSheets);
    formatLaborMgmt(laborSheet);
  }

  // 外注費管理
  var vendorSheet = ss.getSheetByName("外注費管理");
  if (vendorSheet) {
    buildVendorMgmt(vendorSheet, pjList, plSheets);
    formatVendorMgmt(vendorSheet);
  }

  // PJ台帳
  var ledgerSheet = ss.getSheetByName("PJ台帳");
  if (ledgerSheet) {
    buildLedger(ledgerSheet, pjList);
    formatLedger(ledgerSheet);
  }

  ss.toast("完了しました！", "PJ PL管理", 5);
  ss.setActiveSheet(ss.getSheetByName("ダッシュボード"));
}

// ─── 個別PLのみ再構築 ───
function rebuildPLSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pjList = getPjList();
  var plSheets = getPlSheetsWithDefaults();

  ss.toast("PLシートを再構築中...", "PJ PL管理", -1);
  pjList.forEach(function(pj) {
    var title = plSheetTitle(pj.name);
    var sheet = ss.getSheetByName(title);
    if (sheet) {
      sheet.clearContents();
      sheet.clearFormats();
      buildPlSheet(sheet, pj, plSheets[pj.key]);
      formatPlSheet(sheet, plSheets[pj.key].monthLabels.length);
    }
  });

  ss.toast("PLシートの再構築が完了しました！", "PJ PL管理", 5);
}

// ─── ダッシュボードのみ更新 ───
function rebuildDashboardOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pjList = getPjList();
  var plSheets = getPlSheetsWithDefaults();

  ss.toast("ダッシュボードを更新中...", "PJ PL管理", -1);

  var sheets = {
    "ダッシュボード": function(s) { buildDashboard(s, pjList, plSheets); formatDashboard(s, pjList.length); },
    "工数管理": function(s) { buildLaborMgmt(s, pjList, plSheets); formatLaborMgmt(s); },
    "外注費管理": function(s) { buildVendorMgmt(s, pjList, plSheets); formatVendorMgmt(s); },
    "PJ台帳": function(s) { buildLedger(s, pjList); formatLedger(s); },
  };

  Object.keys(sheets).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.clearContents();
      sheet.clearFormats();
      sheets[name](sheet);
    }
  });

  ss.toast("ダッシュボードの更新が完了しました！", "PJ PL管理", 5);
}

// ─── 内部関数: シート削除と再作成 ───
function deleteAndRecreateSheets_(ss, pjList) {
  // 作成するシート一覧
  var sheetNames = ["ダッシュボード"];
  pjList.forEach(function(pj) {
    sheetNames.push(plSheetTitle(pj.name));
  });
  sheetNames.push("工数管理");
  sheetNames.push("外注費管理");
  sheetNames.push("PJ台帳");

  // 一時シートを作成（全削除するとエラーになるため）
  var tmpSheet = ss.insertSheet("_tmp_rebuild_");

  // 既存シートを全削除
  var existingSheets = ss.getSheets();
  existingSheets.forEach(function(s) {
    if (s.getName() !== "_tmp_rebuild_") {
      ss.deleteSheet(s);
    }
  });

  // 新しいシートを作成
  sheetNames.forEach(function(name) {
    ss.insertSheet(name);
  });

  // 一時シートを削除
  ss.deleteSheet(tmpSheet);
}
