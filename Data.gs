/**
 * PJ PL管理 — データ定義
 * PJの追加・変更はこのファイルだけ編集する
 */

// ─── 固定レイアウト定数（0-indexed） ───
var R_TITLE = 0;
var R_INFO = 1;
var R_EMPTY1 = 4;
var R_SUMMARY_H = 5;
var R_SUMMARY_V = 6;
var R_EMPTY2 = 7;
var R_MONTHLY_H = 8;
var R_SALES = 9;
var R_COST = 10;
var R_INTERNAL = 11;
var R_EXTERNAL = 12;
var R_GROSS = 13;
var R_GROSS_PCT = 14;
var R_EMPTY3 = 15;
var R_LABOR_H = 16;
var R_LABOR_START = 17;
var MAX_MEMBERS = 5;
var R_LABOR_SUB = R_LABOR_START + MAX_MEMBERS; // 22
var R_EMPTY4 = R_LABOR_SUB + 1; // 23
var R_VENDOR_H = R_EMPTY4 + 1;  // 24
var R_VENDOR_START = R_VENDOR_H + 1; // 25
var MAX_VENDORS = 5;
var R_VENDOR_SUB = R_VENDOR_START + MAX_VENDORS; // 30

var UNIT_PRICE = 800000;
var C_START = 2;  // PLセクション月データ開始列（C列）
var L_START = 3;  // 工数セクション月データ開始列（D列、C列は単価）

// ─── 色定数（16進数） ───
var CLR_DARK_BLUE = "#264073";
var CLR_LIGHT_BLUE = "#D9EBFF";
var CLR_LIGHT_GRAY = "#F2F2F2";
var CLR_LIGHT_GREEN = "#D9F2D9";
var CLR_LIGHT_YELLOW = "#FFFAD9";
var CLR_MID_BLUE = "#99BFF2";
var CLR_LIGHT_RED = "#FFEBEB";
var CLR_WHITE = "#FFFFFF";

// ─── PJデータ ───
function getPjList() {
  return [
    { code: "PJ0001", name: "TMC（WBT）", company: "ウーブン・バイ・トヨタ株式会社", type: "コンサル", amount: 9475000, start: "2026/01/05", end: "2026/03/13", months: 3, monthly: 3158333, key: "WBT" },
    { code: "PJ0002", name: "ASAP", company: "ASAP SECURITY 株式会社", type: "システム開発", amount: 2800000, start: "2025/11/01", end: "2026/02/28", months: 4, monthly: 700000, key: "ASAP" },
    { code: "PJ0005", name: "清田軌道", company: "清田軌道工業（株）", type: "研修", amount: 4550000, start: "2026/01/27", end: "2026/01/28", months: 1, monthly: 4550000, key: "清田軌道" },
    { code: "PJ0006", name: "ケイズ", company: "株式会社ケイズグループ", type: "その他", amount: 3500000, start: "2025/10/01", end: "2026/04/30", months: 7, monthly: 500000, key: "ケイズ" },
    { code: "PJ0018", name: "Massive", company: "株式会社MASSIVE SAPPORO", type: "システム開発", amount: 14800000, start: "2026/01/01", end: "2026/07/30", months: 7, monthly: 2114286, key: "Massive" },
    { code: "PJ0004", name: "ユーグレナ", company: "株式会社ユーグレナ", type: "提携", amount: 0, start: "", end: "", months: 1, monthly: 0, key: "ユーグレナ" },
    { code: "PJ0007", name: "栗田工業（IST）", company: "栗田工業株式会社", type: "研修", amount: 3900000, start: "2025/11/01", end: "2026/02/28", months: 4, monthly: 975000, key: "栗田工業" },
    { code: "PJ0008", name: "小野薬品（AIMG）", company: "小野薬品工業 株式会社", type: "研修", amount: 1200000, start: "2026/02/01", end: "2026/02/28", months: 1, monthly: 1200000, key: "小野薬品" },
    { code: "PJ0011", name: "パナソニック", company: "パナソニック株式会社", type: "研修", amount: 3000000, start: "2026/02/01", end: "2026/03/31", months: 2, monthly: 1500000, key: "パナソニック" },
    { code: "PJ0017", name: "ONE COMPATH", company: "株式会社ONE COMPATH", type: "システム開発", amount: 3000000, start: "2026/01/01", end: "2026/03/31", months: 3, monthly: 1000000, key: "ONE_COMPATH" },
  ];
}

// ─── PL詳細データ ───
function getPlSheets() {
  return {
    "WBT": {
      monthLabels: ["2026/1", "2026/2", "2026/3"],
      sales: [3158333, 3158333, 3158333],
      members: [
        { name: "斉藤", pct: [50, 50, 50] },
        { name: "笠井", pct: [10, 10, 10] },
        { name: "一谷", pct: [10, 10, 10] },
      ],
      vendors: [],
      status: { "締結": false, "発注書": false, "請求書": false, "計上": false },
      team: "PM: 斉藤  エンジニア: 笠井, 一谷",
    },
    "ASAP": {
      monthLabels: ["2025/12", "2026/1"],
      sales: [1400000, 1400000],
      members: [{ name: "（社内）", pct: [150, 150] }],
      vendors: [],
      status: { "締結": true, "発注書": true, "請求書": true, "計上": true },
      team: "",
    },
    "清田軌道": {
      monthLabels: ["2026/1"],
      sales: [4550000],
      members: [{ name: "（社内）", pct: [50] }],
      vendors: [],
      status: { "締結": false, "発注書": false, "請求書": false, "計上": false },
      team: "",
    },
    "ケイズ": {
      monthLabels: ["2025/10", "2025/11", "2025/12", "2026/1", "2026/2", "2026/3", "2026/4"],
      sales: [500000, 500000, 500000, 500000, 500000, 500000, 500000],
      members: [{ name: "（社内）", pct: [100, 100, 100, 100, 100, 100, 100] }],
      vendors: [],
      status: { "締結": false, "発注書": false, "請求書": false, "計上": false },
      team: "",
    },
    "Massive": {
      monthLabels: ["2026/2", "2026/3", "2026/4", "2026/5", "2026/6", "2026/7"],
      sales: [2114286, 2114286, 2114286, 2114286, 2114286, 2114286],
      members: [{ name: "佐藤", pct: [0, 20, 20, 20, 20, 20] }],
      vendors: [
        { name: "斉藤", amounts: [550000, 500000, 500000, 500000, 500000, 550000] },
        { name: "岩本", amounts: [0, 500000, 500000, 500000, 500000, 500000] },
        { name: "加藤", amounts: [0, 480000, 480000, 480000, 480000, 480000] },
      ],
      status: { "締結": false, "発注書": false, "請求書": false, "計上": false },
      team: "PM: 斉藤  エンジニア: 佐藤, 加藤, 岩本",
    },
    "ユーグレナ": {
      monthLabels: ["2026/1", "2026/2", "2026/3"],
      sales: [0, 0, 0],
      members: [],
      vendors: [],
      status: { "締結": false, "発注書": false, "請求書": false, "計上": false },
      team: "",
    },
  };
}

// PLデータがない案件にデフォルトテンプレートを追加
function getPlSheetsWithDefaults() {
  var plSheets = getPlSheets();
  var pjList = getPjList();
  pjList.forEach(function(pj) {
    if (!(pj.key in plSheets)) {
      var n = Math.max(pj.months, 3);
      var labels = [];
      for (var i = 0; i < n; i++) labels.push("月" + (i + 1));
      plSheets[pj.key] = {
        monthLabels: labels,
        sales: pj.monthly ? new Array(n).fill(pj.monthly) : new Array(n).fill(0),
        members: [],
        vendors: [],
        status: { "締結": false, "発注書": false, "請求書": false, "計上": false },
        team: "",
      };
    }
  });
  return plSheets;
}

// ─── ヘルパー関数 ───

// 0-indexed列番号をアルファベットに変換（A=0, B=1, ...）
function colLetter(idx) {
  if (idx < 26) return String.fromCharCode(65 + idx);
  return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
}

// 0-indexed → 1-indexed（数式用）
function R(row0) {
  return row0 + 1;
}

// シート名からPLシートタイトルを生成
function plSheetTitle(pjName) {
  return "PL_" + pjName;
}
