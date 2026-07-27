/***********************************************************************
 *  KIWAN'S FINANCE TRACKER — Apps Script backend
 *
 *  Single source of truth: the "Cash" and "Dollar" sheets (the raw
 *  ledger). Everything else (Available Balance, Net Worth, Loans,
 *  Dollar price, Avg spend) is read from clearly-labeled cells on the
 *  "Finance" / "Net worth" / "Avg" / "Loans" sheets, located by
 *  searching for their label text rather than by fixed cell address,
 *  so the API keeps working if rows/columns shift.
 *
 *  This file is the ONLY web app entry point for this project.
 *  There must be exactly one doGet() and one doPost() in the whole
 *  Apps Script project — a second copy (there used to be one in
 *  FinanceEntry.gs) silently shadows this one depending on file load
 *  order, which is what caused the API to serve inconsistent/incorrect
 *  data. Do not add another doGet/doPost anywhere else.
 *
 *  SECRETS: Gemini API key lives in Script Properties, never in code.
 *    Apps Script editor → Project Settings → Script properties → Add:
 *      GEMINI_API_KEY = <your real key>
 *    Optional overrides (all have safe defaults if unset):
 *      GEMINI_MODEL, MONTHLY_LIMIT
 ***********************************************************************/

// ============================================================
// SCRIPT PROPERTIES / CONFIG  (no secrets or financial values hardcoded)
// ============================================================
function getProp_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === undefined || v === "") ? fallback : v;
}
function getGeminiKey_()   { return getProp_("GEMINI_API_KEY", ""); }
function getGeminiModel_() { return getProp_("GEMINI_MODEL", "gemini-2.5-flash"); }
function getMonthlyLimit_() { return Number(getProp_("MONTHLY_LIMIT", 15000)) || 15000; }

/*
 * ACCOUNT NAME MAP — the ONE place that maps every account-name variant
 * found in the Cash sheet to its display key. NBE and Insta are the
 * same physical account (NBE Instapay) and must always merge to "Insta".
 */
var ACCOUNT_ALIASES = {
  "nbe":        "Insta",
  "nbe/insta":  "Insta",
  "insta":      "Insta",
  "hsbc":       "HSBC",
  "cash":       "Cash",
  "vf":         "VF",
  "vf visa":    "VF",
  "telda":      "Telda",
  "qnb":        "Qnb",
  "visa":       "Visa"
};

/*
 * Categories/types that represent internal money movement (transfers
 * between own accounts, loan principal draws/repayments, currency
 * conversions) rather than real income or spending. These are EXCLUDED
 * from income/expense totals but still applied to account balances
 * (computeAccountBalances_ never consults this set).
 */
var SKIP_TYPES = [
  "money movement", "deposit", "refund",
  "insta withdrawal", "dollar withdrawal", "dollar deposit",
  "loan", "repaid",
  "transfer=>insta", "transfer=>vf", "transfer=>telda",
  "transfer=>cash", "transfer=>qnb", "transfer=>visa-misr"
];
var SKIP_TYPES_SET = {};
SKIP_TYPES.forEach(function(k){ SKIP_TYPES_SET[k] = 1; });

/*
 * Types counted as income when the amount is positive. Includes known
 * spelling variants seen in the sheet ("Parents Support" vs "Parental
 * Support") so a typo in a single entry doesn't silently drop it from
 * the total the way a strict-equality SUMIF would.
 */
var INCOME_TYPES = [
  "income", "paycheck", "indriver", "side income",
  "bonus", "salary", "trading", "payback"
];
var INCOME_TYPES_SET = {};
INCOME_TYPES.forEach(function(k){ INCOME_TYPES_SET[k] = 1; });

function isSkipType_(s) {
  return !!SKIP_TYPES_SET[s] || s.indexOf("transfer=>") !== -1;
}
function isIncomeType_(s) {
  if (!s) return false;
  if (INCOME_TYPES_SET[s]) return true;
  // "parental support" / "parents support" / "parent support" ...
  return /^parent/.test(s) && s.indexOf("support") !== -1;
}

// ============================================================
// SHEET COLUMN MAPS — header-detected, not fixed column numbers
// ============================================================
function getCashColMap_(sh) {
  var headers = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {
    month:1, date:2, spent:3, account:4, description:5,
    priority:6, expCat:7, expType:8, mainCat:9, notes:10
  };
  var nameMap = {
    "month":"month", "date":"date", "spent":"spent", "account":"account",
    "description":"description", "priority":"priority",
    "expenses category":"expCat", "expense type":"expType",
    "main category":"mainCat", "notes":"notes"
  };
  headers.forEach(function(h, i) {
    var key = nameMap[String(h || "").trim().toLowerCase()];
    if (key) map[key] = i + 1;
  });
  return map;
}

function getDollarColMap_(sh) {
  var headers = sh.getRange(2, 1, 1, Math.min(sh.getLastColumn(), 10)).getValues()[0];
  var map = { month:1, date:2, spent:3, description:4, dollarType:5 };
  var nameMap = {
    "month":"month", "date":"date", "spent":"spent",
    "description":"description", "transfer/profit":"dollarType", "type":"dollarType"
  };
  headers.forEach(function(h, i) {
    var key = nameMap[String(h || "").trim().toLowerCase()];
    if (key) map[key] = i + 1;
  });
  return map;
}

// ============================================================
// LABEL-BASED CELL LOOKUP — for summary sheets (Finance, Net worth)
// whose figures live in scattered, hand-formatted cells rather than a
// tidy header row. We search by label text instead of a fixed address
// so the API survives the sheet being reorganized.
// ============================================================
function findLabeledValue_(sheet, labels, maxRows, maxCols) {
  if (!sheet) return null;
  var lr = Math.min(sheet.getLastRow(), maxRows || 60);
  var lc = Math.min(sheet.getLastColumn(), maxCols || 40);
  if (lr < 1 || lc < 1) return null;
  var vals = sheet.getRange(1, 1, lr, lc).getValues();
  var wanted = labels.map(function(l) { return l.toLowerCase(); });
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var cell = String(vals[r][c] || "").trim().toLowerCase();
      if (!cell || wanted.indexOf(cell) === -1) continue;
      for (var c2 = c + 1; c2 < vals[r].length; c2++) {
        var v = vals[r][c2];
        if (v !== "" && v !== null) return v;
      }
      if (r + 1 < vals.length) {
        var below = vals[r + 1][c];
        if (below !== "" && below !== null) return below;
      }
    }
  }
  return null;
}

/* Some Finance-sheet sections are a row of side-by-side header cells
 * (e.g. "NBE/Insta | HSBC | Cash") with the values directly beneath
 * each header on the next row, not to the right of it. findLabeledValue_
 * would grab the neighboring box's header text by mistake in that
 * layout (it looks right before it looks down), so this looks strictly
 * one row straight down from the matching label. Label text is
 * normalized (case, whitespace, slashes stripped) so "NBE/ INSTA",
 * "NBE/Insta", "NBE / INSTA" etc. all match the same way. */
function findValueBelowLabel_(sheet, labels, maxRows, maxCols) {
  if (!sheet) return null;
  var lr = Math.min(sheet.getLastRow(), maxRows || 60);
  var lc = Math.min(sheet.getLastColumn(), maxCols || 60);
  if (lr < 2 || lc < 1) return null;
  var vals = sheet.getRange(1, 1, lr, lc).getValues();
  var norm = function(s) { return String(s || "").toLowerCase().replace(/[\s\/]/g, ""); };
  var wanted = labels.map(norm);
  for (var r = 0; r < vals.length - 1; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var cell = norm(vals[r][c]);
      if (!cell || wanted.indexOf(cell) === -1) continue;
      var below = vals[r + 1][c];
      if (below !== "" && below !== null) return below;
    }
  }
  return null;
}

/* NBE/Insta running balance, read straight from the Finance sheet's
 * "ACCOUNT BALANCES" box (Finance!AA19 as of this writing) instead of
 * the bottom-up Cash-ledger sum. The ledger sum tracked this cell to
 * within ~0.05% (most likely the Finance tab's snapshot lagging the
 * very latest Cash edits by a few minutes) but the user wants this
 * specific cell mirrored exactly, so it takes priority. Located by the
 * "NBE/Insta" label directly above it first (survives the sheet being
 * reorganized); falls back to a literal AA19 read, then to the ledger
 * sum if neither is available. */
function getInstaBalanceOverride_(ss) {
  var fin = ss.getSheetByName("Finance");
  if (!fin) return null;
  var v = findValueBelowLabel_(fin, ["NBE/Insta", "NBE / INSTA", "NBE/ INSTA"]);
  if (v === null || v === "") {
    try { v = fin.getRange("AA19").getValue(); } catch (e) { v = null; }
  }
  var n = parseMoneyCell_(v);
  return isNaN(n) ? null : n;
}

/* Parses a cell that may be a live number OR currency-formatted text
 * (e.g. "118,927.88 EGP") into a plain number. */
function parseMoneyCell_(v) {
  if (typeof v === "number") return v;
  if (v === null || v === undefined || v === "") return NaN;
  var cleaned = String(v).replace(/[^0-9.\-]/g, "");
  return cleaned ? Number(cleaned) : NaN;
}

// ============================================================
// CENTRALIZED BALANCE / TOTALS CALCULATIONS (Cash sheet = truth)
// ============================================================

/* Sums ALL rows across all months, merging account-name variants via
 * ACCOUNT_ALIASES (nbe/insta -> Insta, etc). This is the ONLY place
 * account balances are computed. */
function computeAccountBalances_(ss) {
  var sh = ss.getSheetByName("Cash");
  if (!sh) { Logger.log("computeAccountBalances_: no Cash sheet"); return {}; }
  var col = getCashColMap_(sh);
  var lr = sh.getLastRow();
  if (lr < 3) return {};

  var totalCols = Math.max(col.month, col.account, col.spent);
  var vals = sh.getRange(3, 1, lr - 2, totalCols).getValues();

  var balances = {};
  vals.forEach(function(r) {
    var amt = Number(r[col.spent - 1]) || 0;
    if (amt === 0) return;
    var rawAcc = String(r[col.account - 1] || "").trim();
    if (!rawAcc) return;
    var key = ACCOUNT_ALIASES[rawAcc.toLowerCase()] || rawAcc;
    balances[key] = (balances[key] || 0) + amt;
  });
  return balances;
}

/* Income / expense / saved for a given month name, excluding internal
 * transfers and loan principal movement (see SKIP_TYPES/INCOME_TYPES). */
function computeMonthTotals_(ss, thisMonth) {
  var sh = ss.getSheetByName("Cash");
  if (!sh) return { income:0, expense:0, saved:0 };
  var col = getCashColMap_(sh);
  var lr = sh.getLastRow();
  if (lr < 3) return { income:0, expense:0, saved:0 };

  var totalCols = Math.max(col.expCat, col.expType, col.mainCat, col.spent, col.month);
  var vals = sh.getRange(3, 1, lr - 2, totalCols).getValues();

  var income = 0, expense = 0;
  vals.forEach(function(r) {
    var m = String(r[col.month - 1] || "").trim();
    if (m.toLowerCase() !== thisMonth.toLowerCase()) return;
    var expCat  = String(r[col.expCat  - 1] || "").trim().toLowerCase();
    var expType = String(r[col.expType - 1] || "").trim().toLowerCase();
    var mainCat = String(r[col.mainCat - 1] || "").trim().toLowerCase();
    var amt = Number(r[col.spent - 1]) || 0;
    if (amt === 0) return;
    if (isSkipType_(expCat) || isSkipType_(expType)) return;
    if (isIncomeType_(expCat) || isIncomeType_(expType) || isIncomeType_(mainCat)) {
      if (amt > 0) income += amt;
    } else if (amt < 0) {
      expense += Math.abs(amt);
    }
  });
  return { income:income, expense:expense, saved:income - expense };
}

/* Amount spent today (excludes internal transfers/loans, same rule as
 * computeMonthTotals_). */
function computeSpentToday_(ss) {
  var sh = ss.getSheetByName("Cash");
  if (!sh) return 0;
  var col = getCashColMap_(sh);
  var lr = sh.getLastRow();
  if (lr < 3) return 0;
  var tz = ss.getSpreadsheetTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, "d/M/yyyy");
  var totalCols = Math.max(col.date, col.spent, col.expCat, col.expType);
  var vals = sh.getRange(3, 1, lr - 2, totalCols).getValues();
  var spent = 0;
  vals.forEach(function(r) {
    var d = r[col.date - 1];
    var ds = (d instanceof Date) ? Utilities.formatDate(d, tz, "d/M/yyyy") : String(d || "").trim();
    var amt = Number(r[col.spent - 1]) || 0;
    var expCat  = String(r[col.expCat  - 1] || "").trim().toLowerCase();
    var expType = String(r[col.expType - 1] || "").trim().toLowerCase();
    if (ds === todayStr && amt < 0 && !isSkipType_(expCat) && !isSkipType_(expType)) {
      spent += Math.abs(amt);
    }
  });
  return spent;
}

/* Average monthly spend from the Avg sheet's Q3:Q11 block (a
 * hand-built summary column, not a header-based table, so it can't be
 * header-detected the way Cash/Dollar are). Falls back to the Finance
 * sheet's "AVG MONTHLY SPEND" label if the Avg sheet is missing. */
function computeAvgSpend_(ss) {
  var a = ss.getSheetByName("Avg");
  if (a) {
    try {
      var vals = a.getRange("Q3:Q11").getValues();
      var total = 0;
      vals.forEach(function(r) {
        var v = Math.abs(Number(r[0]) || 0);
        if (v > 0) total += v;
      });
      if (total > 0) return total;
    } catch (e) {
      Logger.log("computeAvgSpend_: Avg!Q3:Q11 read failed: " + e);
    }
  }
  var fin = ss.getSheetByName("Finance");
  var v2 = findLabeledValue_(fin, ["avg monthly spend", "average monthly spending"]);
  return Math.abs(Number(v2) || 0);
}

/* Per-main-category breakdown with detail sub-items, same skip/income
 * rules as computeMonthTotals_. */
function computeMonthBreakdown_(ss, thisMonth) {
  var sh = ss.getSheetByName("Cash");
  if (!sh) return { income:{}, expense:{} };
  var col = getCashColMap_(sh);
  var lr = sh.getLastRow();
  if (lr < 3) return { income:{}, expense:{} };

  var totalCols = Math.max(col.expCat, col.expType, col.mainCat, col.spent, col.month);
  var vals = sh.getRange(3, 1, lr - 2, totalCols).getValues();

  var income = {}, expense = {};
  vals.forEach(function(r) {
    var m = String(r[col.month - 1] || "").trim();
    if (m.toLowerCase() !== thisMonth.toLowerCase()) return;
    var expCat  = String(r[col.expCat  - 1] || "").trim();
    var expType = String(r[col.expType - 1] || "").trim();
    var mainCat = String(r[col.mainCat - 1] || "").trim() || "Other";
    var amt = Number(r[col.spent - 1]) || 0;
    if (amt === 0) return;
    var catL = expCat.toLowerCase(), typL = expType.toLowerCase(), mainL = mainCat.toLowerCase();
    if (isSkipType_(catL) || isSkipType_(typL)) return;
    var detail = expType || expCat || "Other";
    if (isIncomeType_(catL) || isIncomeType_(typL) || isIncomeType_(mainL)) {
      if (amt > 0) {
        if (!income[mainCat]) income[mainCat] = { total:0, items:{} };
        income[mainCat].total += amt;
        income[mainCat].items[detail] = (income[mainCat].items[detail] || 0) + amt;
      }
    } else if (amt < 0) {
      if (!expense[mainCat]) expense[mainCat] = { total:0, items:{} };
      expense[mainCat].total += Math.abs(amt);
      expense[mainCat].items[detail] = (expense[mainCat].items[detail] || 0) + Math.abs(amt);
    }
  });
  return { income:income, expense:expense };
}

function getAvailableMonths_(ss) {
  var sh = ss.getSheetByName("Cash");
  if (!sh) return [];
  var col = getCashColMap_(sh);
  var lr = sh.getLastRow();
  if (lr < 3) return [];
  var vals = sh.getRange(3, 1, lr - 2, col.month).getValues();
  var seen = {}, months = [];
  vals.forEach(function(r) {
    var m = String(r[col.month - 1] || "").trim();
    if (m && !seen[m]) { seen[m] = true; months.push(m); }
  });
  return months;
}

/* Sum of the Dollar sheet's "spent" column = current dollar holdings. */
function computeDollarHoldings_(ss) {
  var sh = ss.getSheetByName("Dollar");
  if (!sh) return 0;
  var col = getDollarColMap_(sh);
  var lr = sh.getLastRow();
  if (lr < 3) return 0;
  var vals = sh.getRange(3, col.spent, lr - 2, 1).getValues();
  var total = 0;
  vals.forEach(function(r) { total += Number(r[0]) || 0; });
  return total;
}

/* EGP-per-USD rate. No hardcoded fallback number — a stale hardcoded
 * rate is exactly the kind of "hardcoded financial value" this project
 * was full of. If the sheet doesn't expose one, we return 0 and log it
 * rather than guess. */
function getDollarPrice_(ss) {
  var fin = ss.getSheetByName("Finance");
  var v = findLabeledValue_(fin, ["dollar price", "dollar"]);
  if (v && Number(v) > 0) return Number(v);
  var nw = ss.getSheetByName("Net worth");
  var v2 = findLabeledValue_(nw, ["dollar price", "live dollar", "live dollar price"]);
  if (v2 && Number(v2) > 0) return Number(v2);
  Logger.log("getDollarPrice_: no dollar price found on Finance or Net worth sheets");
  return 0;
}

/* Net worth + assets. Prefers the "Net worth" sheet's labeled rows
 * (existing convention: "Total Wealth" / "Assets"); falls back to the
 * Finance sheet's own "NET WORTH" / "Assets Value" cells. */
function computeNetWorthAssets_(ss) {
  var netWorth = 0, assetsVal = 0;
  var nw = ss.getSheetByName("Net worth");
  if (nw) {
    var lr = nw.getLastRow();
    if (lr > 0) {
      var vals = nw.getRange(1, 1, lr, 2).getValues();
      vals.forEach(function(r) {
        var label = String(r[0] || "").trim().toLowerCase();
        if (label === "total wealth") netWorth = Number(r[1]) || 0;
        if (label === "assets")       assetsVal = Number(r[1]) || 0;
      });
    }
  }
  if (!netWorth) {
    var fin = ss.getSheetByName("Finance");
    var v = findLabeledValue_(fin, ["net worth"]);
    if (v) netWorth = Number(v) || 0;
    if (!assetsVal) {
      var v2 = findLabeledValue_(fin, ["assets value", "assets"]);
      if (v2) assetsVal = Number(v2) || 0;
    }
  }
  return { netWorth:netWorth, assetsVal:assetsVal };
}

/* Net loans/debt figure. Prefers the Finance sheet's "Loans/Debt"
 * summary cell; falls back to the Loans sheet's own Grand Total row
 * (EGP column) if the Finance sheet doesn't have it. Negative = money
 * owed TO Kiwan; positive = Kiwan owes it. */
function computeLoansTotal_(ss) {
  var fin = ss.getSheetByName("Finance");
  var v = findLabeledValue_(fin, ["loans/debt", "loans / debt", "loans"]);
  if (v !== null && v !== "") return Number(v) || 0;

  var loanSh = ss.getSheetByName("Loans");
  if (!loanSh) return 0;
  var lr = loanSh.getLastRow();
  if (lr < 2) return 0;
  var egp = loanSh.getRange(2, 1, lr - 1, 2).getValues();
  var total = 0, found = false;
  egp.forEach(function(r) {
    var nm = String(r[0] || "").trim().toLowerCase();
    if (nm.indexOf("grand total") !== -1) { total = Number(r[1]) || 0; found = true; }
  });
  if (!found) {
    Logger.log("computeLoansTotal_: no Grand Total row found on Loans sheet");
  }
  return total;
}

// ============================================================
// doGet: balances + AI ask (JSONP) — the ONLY doGet in this project
// ============================================================
function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "cb";
  try {
    return handleGet_(e, cb);
  } catch (err) {
    Logger.log("doGet error: " + err + (err && err.stack ? "\n" + err.stack : ""));
    var errOut = { status:"error", message: String(err) };
    return ContentService
      .createTextOutput(cb + "(" + JSON.stringify(errOut) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function handleGet_(e, cb) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : null;
  var out;

  if (action === "ask") {
    out = { status:"ok", answer: askGemini_(ss, e.parameter.q || "") };

  } else if (action === "recent") {
    out = recentEntries_(ss);

  } else if (action === "voice") {
    out = voiceEntry_(ss, e.parameter.text || "");

  } else if (action === "memory") {
    out = buildMemory_(ss);

  } else if (action === "update") {
    out = updateRow_(ss, e.parameter);

  } else if (action === "delete") {
    out = deleteRow_(ss, e.parameter.sheet, Number(e.parameter.row));

  } else {
    // Default: balances + month summary. Accepts optional ?month=June.
    var tz = ss.getSpreadsheetTimeZone();
    var reqMonth = (e && e.parameter && e.parameter.month) ? e.parameter.month : null;
    var dashSh = ss.getSheetByName("Dashboard");
    var dashMonth = dashSh ? String(dashSh.getRange("A7").getValue() || "").trim() : "";

    var thisMonth;
    if (reqMonth) {
      // The app explicitly asked for a month (user tapped a month chip)
      // -> that wins, and we push it back so Dashboard!A7 matches the
      // app's current view.
      thisMonth = reqMonth;
      if (dashSh) dashSh.getRange("A7").setValue(reqMonth);
    } else if (dashMonth) {
      // No month requested -> follow whatever Dashboard!A7 already says,
      // so opening the app always reflects the sheet's selected month.
      thisMonth = dashMonth;
    } else {
      thisMonth = Utilities.formatDate(new Date(), tz, "MMMM");
    }

    var balances = computeAccountBalances_(ss);
    var instaOverride = getInstaBalanceOverride_(ss);
    if (instaOverride !== null) {
      Logger.log("handleGet_: overriding Insta balance with Finance-sheet value " +
        instaOverride + " (ledger sum was " + (balances["Insta"] || 0) + ")");
      balances["Insta"] = instaOverride;
    }

    var totals         = computeMonthTotals_(ss, thisMonth);
    var breakdown      = computeMonthBreakdown_(ss, thisMonth);
    var avgSpend       = computeAvgSpend_(ss);
    var allMonths      = getAvailableMonths_(ss);
    var dollarHoldings = computeDollarHoldings_(ss);
    var dollarPrice    = getDollarPrice_(ss);
    var netWorthAssets = computeNetWorthAssets_(ss);
    var loans          = computeLoansTotal_(ss);

    var accountsSum = 0;
    Object.keys(balances).forEach(function(k) { accountsSum += balances[k]; });

    var fin = ss.getSheetByName("Finance");
    var availBalRaw = findLabeledValue_(fin, ["available balance"]);
    var availableBalance = (availBalRaw !== null && availBalRaw !== "")
      ? Number(availBalRaw) || 0
      : (accountsSum + dollarHoldings * dollarPrice);

    Logger.log("doGet default: month=" + thisMonth +
      " income=" + totals.income + " expense=" + totals.expense +
      " balances=" + JSON.stringify(balances));

    out = {
      // Account balances
      insta:    balances["Insta"]  || 0,
      hsbc:     balances["HSBC"]   || 0,
      cash:     balances["Cash"]   || 0,
      vf:       balances["VF"]     || 0,
      visamisr: balances["Telda"]  || 0,
      qnb:      balances["Qnb"]    || 0,
      visa:     balances["Visa"]   || 0,
      // Monthly summary
      income:   totals.income,
      expense:  totals.expense,
      saved:    totals.saved,
      // Breakdown by type (for tap-to-expand In/Out)
      income_breakdown:  breakdown.income,
      expense_breakdown: breakdown.expense,
      // Month navigation
      current_month:     thisMonth,
      available_months:  allMonths,
      // Spend tracking
      avgSpend:    getMonthlyLimit_(),
      avgSpending: avgSpend,
      spentToday:  computeSpentToday_(ss),
      // Wealth summary
      available_balance: availableBalance,
      net_worth:          netWorthAssets.netWorth,
      assets_val:         netWorthAssets.assetsVal,
      loans:              loans,
      dollar_holdings:    dollarHoldings,
      dollar_price:       dollarPrice
    };
  }

  return ContentService
    .createTextOutput(cb + "(" + JSON.stringify(out) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ============================================================
// doPost: add entry — the ONLY doPost in this project
// ============================================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonFE_({ status:"error", message:"Missing request body." });
    }
    var data = JSON.parse(e.postData.contents);
    if (!data.sheet || (data.sheet !== "Cash" && data.sheet !== "Dollar")) {
      return jsonFE_({ status:"error", message:"data.sheet must be 'Cash' or 'Dollar'." });
    }
    var amount = Number(data.amount);
    if (!amount || isNaN(amount)) {
      return jsonFE_({ status:"error", message:"data.amount must be a non-zero number." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var tz = ss.getSpreadsheetTimeZone();
    var monthName = Utilities.formatDate(now, tz, "MMMM");

    if (data.sheet === "Cash") {
      var sh = ss.getSheetByName("Cash");
      if (!sh) return jsonFE_({ status:"error", message:"Cash sheet not found." });
      var col = getCashColMap_(sh);
      var row = firstEmptyRowFE_(sh, col.month, col.spent, 3);
      sh.getRange(row, col.month).setValue(monthName);
      sh.getRange(row, col.date).setValue(now);
      sh.getRange(row, col.spent).setValue(amount);
      sh.getRange(row, col.account).setValue(data.account || "");
      sh.getRange(row, col.description).setValue(data.description || "");
      sh.getRange(row, col.expType).setValue(data.expenseType || "");

    } else {
      var dsh = ss.getSheetByName("Dollar");
      if (!dsh) return jsonFE_({ status:"error", message:"Dollar sheet not found." });
      var dcol = getDollarColMap_(dsh);
      var drow = firstEmptyRowFE_(dsh, dcol.month, dcol.spent, 3);
      dsh.getRange(drow, dcol.month).setValue(monthName);
      dsh.getRange(drow, dcol.date).setValue(now);
      dsh.getRange(drow, dcol.spent).setValue(amount);
      dsh.getRange(drow, dcol.description).setValue(data.description || "");
      dsh.getRange(drow, dcol.dollarType).setValue(data.dollarType || "");
    }
    return jsonFE_({ status:"ok" });

  } catch (err) {
    Logger.log("doPost error: " + err);
    return jsonFE_({ status:"error", message: String(err) });
  }
}

// ============================================================
// Gemini: AI Ask
// ============================================================
function askGemini_(ss, question) {
  var key = getGeminiKey_();
  if (!key) return "Gemini key not set. Add GEMINI_API_KEY in Project Settings → Script properties.";
  if (!question) return "Ask me something about your finances.";

  var tz = ss.getSpreadsheetTimeZone();
  var thisMonth = Utilities.formatDate(new Date(), tz, "MMMM");

  var totals   = computeMonthTotals_(ss, thisMonth);
  var balances = computeAccountBalances_(ss);
  var income   = totals.income;
  var expense  = totals.expense;
  var saved    = totals.saved;

  var mcSh = ss.getSheetByName("Main Categories");
  var mainCats = [];
  var details = [];
  if (mcSh) {
    var mcLr = mcSh.getLastRow();
    if (mcLr > 2) {
      var mcVals = mcSh.getRange(3, 1, mcLr - 2, 3).getValues();
      mcVals.forEach(function(r) {
        var name = String(r[0] || "").trim();
        var actual = Number(r[1]) || 0;
        if (!name) return;
        if (name.startsWith("*")) {
          if (actual !== 0) details.push(name.replace(/^\*\s*/, "") + ": " + actual);
        } else {
          if (actual !== 0) mainCats.push(name + ": " + actual);
        }
      });
    }
  }

  var cashSh = ss.getSheetByName("Cash");
  var col = cashSh ? getCashColMap_(cashSh) : null;
  var rawLines = [];
  if (cashSh && col) {
    var lr = cashSh.getLastRow();
    if (lr > 2) {
      var totalCols = Math.max(col.expType, col.account, col.description, col.spent, col.date, col.month);
      var vals = cashSh.getRange(3, 1, lr - 2, totalCols).getValues();
      vals.forEach(function(r) {
        var m = String(r[col.month - 1] || "");
        if (m.toLowerCase() !== thisMonth.toLowerCase()) return;
        rawLines.push("Cash | " + r.slice(0, totalCols).join(" | "));
      });
    }
  }
  var dollarSh = ss.getSheetByName("Dollar");
  if (dollarSh) {
    var dcol = getDollarColMap_(dollarSh);
    var dlr = dollarSh.getLastRow();
    if (dlr > 2) {
      var dvals = dollarSh.getRange(3, 1, dlr - 2, Math.max(dcol.month, dcol.spent, dcol.description, dcol.dollarType)).getValues();
      dvals.forEach(function(r) {
        var m = String(r[dcol.month - 1] || "");
        if (m.toLowerCase() !== thisMonth.toLowerCase()) return;
        rawLines.push("Dollar | " + r.join(" | "));
      });
    }
  }

  var loanLines = [];
  var loanSh = ss.getSheetByName("Loans");
  if (loanSh) {
    var llr = loanSh.getLastRow();
    if (llr >= 2) {
      var egp = loanSh.getRange(2, 1, llr - 1, 2).getValues();
      var usd = loanSh.getRange(2, 3, llr - 1, 2).getValues();
      egp.forEach(function(r) {
        var nm = String(r[0] || "").trim();
        var amt = Number(r[1]) || 0;
        if (nm && nm.toLowerCase().indexOf("grand total") === -1 && amt !== 0)
          loanLines.push(nm + ": " + amt + " EGP (" + (amt < 0 ? "they owe Kiwan" : "Kiwan owes") + ")");
      });
      usd.forEach(function(r) {
        var nm = String(r[0] || "").trim();
        var amt = Number(r[1]) || 0;
        if (nm && nm.toLowerCase().indexOf("grand total") === -1 && amt !== 0)
          loanLines.push(nm + ": " + amt + " USD (" + (amt < 0 ? "they owe Kiwan" : "Kiwan owes") + ")");
      });
    }
  }

  function dumpSheet_(name, maxRows, maxCols) {
    var sh = ss.getSheetByName(name);
    if (!sh) return "";
    var lr = Math.min(sh.getLastRow(), maxRows);
    var lc = Math.min(sh.getLastColumn(), maxCols);
    if (lr < 1 || lc < 1) return "";
    var vals = sh.getRange(1, 1, lr, lc).getValues();
    var lines = [];
    vals.forEach(function(row) {
      var cells = row.map(function(c) {
        if (c === null || c === "") return "";
        if (c instanceof Date) return Utilities.formatDate(c, tz, "d/M/yyyy");
        if (typeof c === "number") return Math.round(c * 100) / 100;
        return String(c);
      });
      while (cells.length && String(cells[cells.length - 1]).trim() === "") cells.pop();
      if (cells.some(function(c) { return String(c).trim() !== ""; }))
        lines.push(cells.join(" | "));
    });
    return lines.length ? ("--- TAB: " + name + " ---\n" + lines.join("\n")) : "";
  }

  var avgBlock      = dumpSheet_("Avg", 20, 17);
  var netWorthBlock = dumpSheet_("Net worth", 20, 5);
  var incomeBlock   = dumpSheet_("Income", 20, 12);

  var balStr = Object.keys(balances).map(function(k){ return k + ": " + Math.round(balances[k]); }).join(", ");

  var context =
    "Current month: " + thisMonth + "\n" +
    "All amounts are EGP unless marked USD. Negative = expense/outflow.\n\n" +
    "========== PRIMARY ==========\n" +
    "THIS MONTH:\n" +
    "- Income this month: " + Math.round(income) + "\n" +
    "- Total expenses this month: " + Math.round(expense) + "\n" +
    "- Saved this month: " + Math.round(saved) + "\n\n" +
    "CURRENT ACCOUNT BALANCES (running totals):\n" + balStr + "\n\n" +
    "SPENDING BY MAIN CATEGORY:\n" + mainCats.join("\n") + "\n\n" +
    "SPENDING BY DETAILED CATEGORY (non-zero):\n" + details.join("\n") + "\n\n" +
    "LOANS / DEBTS:\n" + (loanLines.length ? loanLines.join("\n") : "None") + "\n\n" +
    "========== SECONDARY ==========\n" +
    "MONTHLY AVERAGES:\n" + avgBlock + "\n\n" +
    "NET WORTH:\n" + netWorthBlock + "\n\n" +
    "INCOME TAB:\n" + incomeBlock + "\n\n" +
    "========== RAW TRANSACTIONS ==========\n" +
    rawLines.join("\n");

  var prompt =
    "You are Kiwan's personal financial advisor. You have FULL access to his finance workbook. " +
    "Answer in the SAME language as the question (Egyptian Arabic if Arabic). Be warm, specific with real numbers.\n\n" +
    "HOW TO USE:\n" +
    "- PRIMARY section = accurate headline figures.\n" +
    "- SECONDARY section = trends, history, net worth — use for analysis.\n" +
    "- RAW TRANSACTIONS = detailed lookup.\n" +
    "- For debts: negative = person owes Kiwan; positive = Kiwan owes them.\n\n" +
    "YOUR JOB: Act like a real financial advisor. Give analysis, comparisons, concrete plans. " +
    "Be honest about overspending vs averages.\n\n" +
    "DATA:\n" + context + "\n\nQUESTION: " + question;

  var url = "https://generativelanguage.googleapis.com/v1beta/models/" +
            getGeminiModel_() + ":generateContent?key=" + key;
  var payload = { contents: [{ parts: [{ text: prompt }] }] };
  try {
    var res = UrlFetchApp.fetch(url, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify(payload), muteHttpExceptions:true
    });
    var json = JSON.parse(res.getContentText());
    if (json.candidates && json.candidates[0]) {
      return json.candidates[0].content.parts[0].text;
    }
    return "Error: " + (json.error ? json.error.message : "No answer.");
  } catch (err) {
    Logger.log("askGemini_ error: " + err);
    return "Error contacting Gemini: " + String(err);
  }
}

// ============================================================
// VOICE ENTRY (via Gemini)
// ============================================================
function voiceEntry_(ss, text) {
  if (!text) return { status:"error", message:"No speech." };
  var key = getGeminiKey_();
  if (!key) return { status:"error", message:"Gemini key not set. Add GEMINI_API_KEY in Script properties." };

  var accounts = ["NBE", "Insta", "Cash", "VF", "Telda", "HSBC", "Qnb", "Visa"];
  var cashSh = ss.getSheetByName("Cash");
  if (cashSh) {
    var accSet = {};
    var lr = cashSh.getLastRow();
    if (lr > 2) {
      var col = getCashColMap_(cashSh);
      var vals = cashSh.getRange(3, 1, Math.min(lr - 2, 300), col.account).getValues();
      vals.forEach(function(r) {
        var a = String(r[col.account - 1] || "").trim();
        if (a) accSet[a] = true;
      });
      var found = Object.keys(accSet).filter(function(a){ return a.length > 0; });
      if (found.length) accounts = found;
    }
  }

  var cashTypes = "Paycheck, Bonus, Indriver, Parental Support, Side Income, Salary, Indriver fees, " +
    "Loan, Repaid, Dollar Withdrawal, Dollar Deposit, Breakfast, Diet Food, Junk, " +
    "snacks, Fuel, Car repair, Clothes, Personal Items, Accessories, Supplements, " +
    "Water, Education fees, Medicine, Internet, Calls, One-Off, Gameaa, Tips& Gratuities, " +
    "Occasional Supplies, Minor Maintenance, Laundry, Lost, Haircut, Withdrawal fees, " +
    "Fruits, Beverages, Oil change, Gym, Apple storage, Chatgpt, Boxing PT, Boxing Group, " +
    "Youtube, Fuel Reimbursement, Engine Oil Reimbursement, " +
    "Transfer=>Insta, Transfer=>VF, Transfer=>Telda, Transfer=>Cash, Transfer=>Qnb, Uber, " +
    "Bank Fees, Gifts, Delivery, Lending, Payback, Deposit, Insta Withdrawal, Refund";
  var dollarTypes = "Transfer In, Transfer Out, Profit, loss, Loan, Repaid";

  var mem = buildMemory_(ss);
  var memHints = [];
  Object.keys(mem.Cash || {}).forEach(function(k) {
    var t = mem.Cash[k].type;
    if (t) memHints.push('"' + k + '" => ' + t);
  });
  var memStr = memHints.slice(0, 60).join("; ");

  var prompt =
    "You convert a spoken Egyptian-Arabic money note into a JSON ARRAY. Return ONLY raw JSON array, no markdown.\n" +
    "SHEET per entry: mentions دولار/دولر/USDT => \"Dollar\"; otherwise => \"Cash\".\n" +
    "Object schema:\n" +
    "{\"sheet\":\"Cash|Dollar\",\"flow\":\"out|in\",\"amount\":number," +
    "\"account\":\"one of: " + accounts.join(", ") + "\"," +
    "\"expenseType\":\"Cash only\",\"dollarType\":\"Dollar only\",\"description\":\"short Arabic note\"}\n" +
    "Rules:\n" +
    "- amount: POSITIVE. Understand Arabic numbers. flow=out for صرفت/اشتريت/دفعت, in for دخل/قبضت.\n" +
    "- ACCOUNT: انستا/NBE=NBE, كاش/نقدي=Cash, QNB/قطر الوطني=Qnb, فودافون/في اف=VF, تلدا=Telda, HSBC=HSBC. " +
    "Apply one stated account to all entries that don't specify their own. Default=Cash.\n" +
    "- Cash expenseType MUST be one of: " + cashTypes + ".\n" +
    "- Dollar dollarType MUST be one of: " + dollarTypes + ". Dollar has NO account/expenseType.\n" +
    "- CATEGORY: check memory [" + memStr + "]. Then guess by meaning.\n" +
    "- description: the item in Arabic.\n" +
    "Spoken note: \"" + text + "\"";

  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + getGeminiModel_() + ":generateContent?key=" + key;
  try {
    var res = UrlFetchApp.fetch(url, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      muteHttpExceptions:true
    });
    var json = JSON.parse(res.getContentText());
    if (!json.candidates || !json.candidates[0])
      return { status:"error", message: (json.error ? json.error.message : "No answer.") };
    var raw = json.candidates[0].content.parts[0].text;
    raw = raw.replace(/```json/gi, "").replace(/```/g, "").trim();
    var data;
    try { data = JSON.parse(raw); } catch (err) {
      return { status:"error", message:"Couldn't understand. Try again.", raw: raw };
    }
    var arr = Array.isArray(data) ? data : [data];
    arr = arr.filter(function(d) { return d && (Number(d.amount) || 0) !== 0; });
    arr.forEach(function(d) { d.amount = Math.abs(Number(d.amount) || 0); });
    if (!arr.length) return { status:"error", message:"No amount detected." };
    return { status:"ok", parsed: arr };
  } catch (err) {
    Logger.log("voiceEntry_ error: " + err);
    return { status:"error", message:"Error contacting Gemini: " + String(err) };
  }
}

// ============================================================
// SMART MEMORY: learn description -> type/account
// ============================================================
function buildMemory_(ss) {
  function learn(name, startR, isColMapFn) {
    var sh = ss.getSheetByName(name);
    var map = {};
    var descSet = {};
    if (!sh) return { map:{}, descs:[] };
    var col = isColMapFn(sh);
    var lr = sh.getLastRow();
    if (lr < startR) return { map:{}, descs:[] };
    var MAXR = 400;
    var from = Math.max(startR, lr - MAXR + 1);
    var totalCols = (name === "Dollar")
      ? Math.max(col.description, col.dollarType)
      : Math.max(col.description, col.expType, col.account);
    var vals = sh.getRange(from, 1, lr - from + 1, totalCols).getValues();
    vals.forEach(function(r) {
      var desc = String(r[col.description - 1] || "").trim();
      if (!desc) return;
      var key = desc.toLowerCase();
      descSet[desc] = true;
      if (!map[key]) map[key] = { type:{}, account:{} };
      if (name === "Dollar") {
        var typ = String(r[col.dollarType - 1] || "").trim();
        if (typ) map[key].type[typ] = (map[key].type[typ] || 0) + 1;
      } else {
        var acc = String(r[col.account - 1] || "").trim();
        var typ = String(r[col.expType  - 1] || "").trim();
        if (acc) map[key].account[acc] = (map[key].account[acc] || 0) + 1;
        if (typ) map[key].type[typ]    = (map[key].type[typ]    || 0) + 1;
      }
    });
    var out = {};
    Object.keys(map).forEach(function(k) {
      out[k] = { type: topKey(map[k].type), account: topKey(map[k].account) };
    });
    return { map:out, descs: Object.keys(descSet) };
  }
  function topKey(obj) {
    var best = "", n = -1;
    Object.keys(obj).forEach(function(k) { if (obj[k] > n) { n = obj[k]; best = k; } });
    return best;
  }
  var cash   = learn("Cash",   3, getCashColMap_);
  var dollar = learn("Dollar", 3, getDollarColMap_);
  return {
    status:"ok",
    Cash:cash.map,     CashDescs:cash.descs,
    Dollar:dollar.map, DollarDescs:dollar.descs
  };
}

// ============================================================
// RECENT ENTRIES (last 60 days)
// ============================================================
function recentEntries_(ss) {
  var tz = ss.getSpreadsheetTimeZone();
  var now = new Date();

  function dayNum(d) {
    var s = Utilities.formatDate(d, tz, "yyyy-MM-dd").split("-");
    return Math.floor(Date.UTC(+s[0], +s[1]-1, +s[2]) / 86400000);
  }
  var todayNum = dayNum(now);

  function parseCell_(v) {
    if (v instanceof Date) return { ymd: dayNum(v), ts: v.getTime() };
    var s = String(v || "").trim();
    if (!s) return null;
    var datePart = s.split(" ")[0];
    var mMatch = datePart.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (mMatch) {
      var months = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
      var mi = months.indexOf(mMatch[2].toLowerCase());
      if (mi >= 0) {
        var dn = Math.floor(Date.UTC(+mMatch[3], mi, +mMatch[1]) / 86400000);
        return { ymd: dn, ts: Date.UTC(+mMatch[3], mi, +mMatch[1]) };
      }
    }
    var m = datePart.split(/[\/\-.]/);
    if (m.length < 3) return null;
    var day = parseInt(m[0], 10), mon = parseInt(m[1], 10), yr = parseInt(m[2], 10);
    if (!day || !mon || !yr) return null;
    if (yr < 100) yr += 2000;
    var dn = Math.floor(Date.UTC(yr, mon - 1, day) / 86400000);
    return { ymd: dn, ts: Date.UTC(yr, mon - 1, day) };
  }

  function pull(name, startR, colMapFn) {
    var sh = ss.getSheetByName(name);
    var arr = [];
    if (!sh) return arr;
    var col = colMapFn(sh);
    var lr = sh.getLastRow();
    if (lr < startR) return arr;
    var totalCols = (name === "Dollar")
      ? Math.max(col.month, col.date, col.spent, col.description, col.dollarType)
      : Math.max(col.month, col.date, col.spent, col.account, col.description, col.expType);
    var vals = sh.getRange(startR, 1, lr - startR + 1, totalCols).getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      if (String(r[col.month - 1] || "").trim() === "" && !r[col.spent - 1]) continue;
      var pd = parseCell_(r[col.date - 1]);
      if (!pd) continue;
      if (todayNum - pd.ymd > 60 || pd.ymd > todayNum) continue;
      var dispDate = Utilities.formatDate(new Date(pd.ts), tz, "d/M/yyyy");
      var obj = {
        row:    startR + i,
        date:   dispDate,
        ts:     pd.ts,
        amount: Number(r[col.spent - 1]) || 0
      };
      if (name === "Dollar") {
        obj.description = String(r[col.description - 1] || "");
        obj.dollarType  = String(r[col.dollarType  - 1] || "");
      } else {
        obj.account     = String(r[col.account     - 1] || "");
        obj.description = String(r[col.description - 1] || "");
        obj.expenseType = String(r[col.expType     - 1] || "");
      }
      arr.push(obj);
    }
    arr.sort(function(a, b) { return b.ts - a.ts; });
    return arr;
  }

  return {
    status:"ok",
    Cash:   pull("Cash",   3, getCashColMap_),
    Dollar: pull("Dollar", 3, getDollarColMap_)
  };
}

// ============================================================
// UPDATE ROW
// ============================================================
function updateRow_(ss, p) {
  try {
    var name = p.sheet;
    var row = Number(p.row);
    if (name !== "Cash" && name !== "Dollar") return { status:"error", message:"bad sheet" };
    var sh = ss.getSheetByName(name);
    if (!sh || !row || row < 3 || row > sh.getMaxRows()) return { status:"error", message:"bad target" };

    if (name === "Dollar") {
      var col = getDollarColMap_(sh);
      if (p.amount      !== undefined) sh.getRange(row, col.spent).setValue(Number(p.amount));
      if (p.description !== undefined) sh.getRange(row, col.description).setValue(p.description);
      if (p.dollarType  !== undefined) sh.getRange(row, col.dollarType).setValue(p.dollarType);
    } else {
      var ccol = getCashColMap_(sh);
      if (p.amount      !== undefined) sh.getRange(row, ccol.spent).setValue(Number(p.amount));
      if (p.account     !== undefined) sh.getRange(row, ccol.account).setValue(p.account);
      if (p.description !== undefined) sh.getRange(row, ccol.description).setValue(p.description);
      if (p.expenseType !== undefined) sh.getRange(row, ccol.expType).setValue(p.expenseType);
    }
    return { status:"ok" };
  } catch (err) {
    Logger.log("updateRow_ error: " + err);
    return { status:"error", message: String(err) };
  }
}

// ============================================================
// DELETE ROW
// ============================================================
function deleteRow_(ss, name, row) {
  try {
    if (name !== "Cash" && name !== "Dollar") return { status:"error", message:"bad sheet" };
    var sh = ss.getSheetByName(name);
    if (!sh || !row || row < 3 || row > sh.getMaxRows()) return { status:"error", message:"bad target" };
    sh.deleteRow(row);
    return { status:"ok" };
  } catch (err) {
    Logger.log("deleteRow_ error: " + err);
    return { status:"error", message: String(err) };
  }
}

// ============================================================
// SHEET-SIDE AUTOMATION (triggers, menu) — unrelated to the web app,
// runs inside the spreadsheet itself.
// ============================================================
function setup() {
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh  = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const val = String(e.range.getDisplayValue()).trim();

  if (sh.getName() === "Cash") {
    const cmap = getCashColMap_(sh);
    if (col === cmap.spent  && row > 2) autoFillCashMonthDate_(sh, row);
    if (col === cmap.expType && row > 2 && val === "Dollar Deposit") handleCashToDollarSync_(sh, row);
  }
  if (sh.getName() === "Dollar") {
    const dmap = getDollarColMap_(sh);
    if (col === dmap.spent      && row > 2) autoFillCashMonthDate_(sh, row);
    if (col === dmap.dollarType && row > 2 && val === "Transfer Out") handleDollarToCashSync_(sh, row);
  }
  if (sh.getName() === "AnQa" || sh.getName() === "Rawaq") {
    if (col === 3 && row > 2) autoFillCashMonthDate_(sh, row);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard Tools')
    .addItem('Open Details',      'openSelectedCategoryDetails')
    .addItem('Clear Cash Filter', 'clearCashFilter')
    .addToUi();
}

function autoFillCashMonthDate_(sheet, row) {
  const amtCol = (function() {
    try { return getCashColMap_(sheet).spent; } catch(e) { return 3; }
  })();
  const amount = sheet.getRange(row, amtCol).getValue();
  if (amount === "" || amount === null) return;
  const dateCell  = sheet.getRange(row, 2);
  const monthCell = sheet.getRange(row, 1);
  const tz        = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const today     = new Date();
  if (dateCell.isBlank()) {
    dateCell.setValue(today);
    dateCell.setNumberFormat("d/m/yyyy");
  }
  const actualDate = dateCell.getValue() || today;
  if (monthCell.isBlank())
    monthCell.setValue(Utilities.formatDate(new Date(actualDate), tz, "MMMM"));
}

function handleCashToDollarSync_(cashSheet, row) {
  const ss          = cashSheet.getParent();
  const dollarSheet = ss.getSheetByName("Dollar");
  if (!dollarSheet) return;
  const cmap = getCashColMap_(cashSheet);
  const amountEGP   = cashSheet.getRange(row, cmap.spent).getValue();
  const descCell    = cashSheet.getRange(row, cmap.description);
  const description = String(descCell.getDisplayValue()).trim();
  const expenseType = String(cashSheet.getRange(row, cmap.expType).getDisplayValue()).trim();
  if (expenseType !== "Dollar Deposit" || !amountEGP || isNaN(amountEGP) || !description) return;
  autoFillCashMonthDate_(cashSheet, row);
  const finalMonth = cashSheet.getRange(row, cmap.month).getValue();
  const finalDate  = cashSheet.getRange(row, cmap.date).getValue();
  const rate = extractRateFromText_(description);
  if (!rate || rate <= 0) return;
  const usdAmount = Math.abs(Number(amountEGP)) / rate;
  const dmap = getDollarColMap_(dollarSheet);
  let linkedRow = getLinkedRowFromNote_(descCell, "SYNC_DOLLAR_ROW");
  if (!linkedRow || linkedRow < 2 || linkedRow > dollarSheet.getMaxRows()) {
    linkedRow = getFirstEmptyRow_(dollarSheet, dmap.spent, 3);
    setLinkedRowNote_(descCell, "SYNC_DOLLAR_ROW", linkedRow);
  }
  dollarSheet.getRange(linkedRow, dmap.month).setValue(finalMonth);
  dollarSheet.getRange(linkedRow, dmap.date).setValue(finalDate);
  dollarSheet.getRange(linkedRow, dmap.date).setNumberFormat("d/m/yyyy");
  dollarSheet.getRange(linkedRow, dmap.spent).setValue(usdAmount);
  dollarSheet.getRange(linkedRow, dmap.spent).setNumberFormat("0.00");
  dollarSheet.getRange(linkedRow, dmap.description).setValue(description);
  dollarSheet.getRange(linkedRow, dmap.dollarType).setValue("Transfer In");
}

function handleDollarToCashSync_(dollarSheet, row) {
  const ss        = dollarSheet.getParent();
  const cashSheet = ss.getSheetByName("Cash");
  if (!cashSheet) return;
  const dmap = getDollarColMap_(dollarSheet);
  const amountUSD   = dollarSheet.getRange(row, dmap.spent).getValue();
  const descCell    = dollarSheet.getRange(row, dmap.description);
  const description = String(descCell.getDisplayValue()).trim();
  const type        = String(dollarSheet.getRange(row, dmap.dollarType).getDisplayValue()).trim();
  if (type !== "Transfer Out" || !amountUSD || isNaN(amountUSD) || Number(amountUSD) >= 0 || !description) return;
  autoFillCashMonthDate_(dollarSheet, row);
  const finalMonth = dollarSheet.getRange(row, dmap.month).getValue();
  const finalDate  = dollarSheet.getRange(row, dmap.date).getValue();
  const rate = extractRateFromText_(description);
  if (!rate || rate <= 0) return;
  const egpAmount = Math.abs(Number(amountUSD)) * rate;
  const cmap = getCashColMap_(cashSheet);
  let linkedRow = getLinkedRowFromNote_(descCell, "SYNC_CASH_ROW");
  if (!linkedRow || linkedRow < 3 || linkedRow > cashSheet.getMaxRows()) {
    linkedRow = getFirstEmptyRow_(cashSheet, cmap.spent, 3);
    setLinkedRowNote_(descCell, "SYNC_CASH_ROW", linkedRow);
  }
  cashSheet.getRange(linkedRow, cmap.month).setValue(finalMonth);
  cashSheet.getRange(linkedRow, cmap.date).setValue(finalDate);
  cashSheet.getRange(linkedRow, cmap.date).setNumberFormat("d/m/yyyy");
  cashSheet.getRange(linkedRow, cmap.spent).setValue(egpAmount);
  cashSheet.getRange(linkedRow, cmap.spent).setNumberFormat("$#,##0.00;-$#,##0.00");
  cashSheet.getRange(linkedRow, cmap.description).setValue(description);
  cashSheet.getRange(linkedRow, cmap.expType).setValue("Dollar Withdrawal");
}

function openSelectedCategoryDetails() {
  const ss             = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("Dashboard");
  const cashSheet      = ss.getSheetByName("Cash");
  if (!dashboardSheet || !cashSheet) return;
  const activeSheet = ss.getActiveSheet();
  const activeCell  = activeSheet.getActiveCell();
  if (activeSheet.getName() !== "Dashboard") {
    SpreadsheetApp.getUi().alert("افتح Dashboard وحدد category.");
    return;
  }
  const row = activeCell.getRow(), col = activeCell.getColumn();
  const isSmall = (col===2&&row>=28&&row<=100)||(col===8&&row>=28&&row<=35);
  const isMain  = col===8&&row>=40&&row<=80;
  if (!isSmall && !isMain) { SpreadsheetApp.getUi().alert("حدد خلية صح."); return; }
  const selectedValue = cleanCategory_(activeCell.getDisplayValue());
  const month         = normalizeMonth_(dashboardSheet.getRange("A7").getDisplayValue());
  if (!selectedValue || !month) { SpreadsheetApp.getUi().alert("مشكلة في الكاتيجوري أو الشهر."); return; }
  const cmap    = getCashColMap_(cashSheet);
  const lastRow = cashSheet.getLastRow();
  const existing = cashSheet.getFilter();
  if (existing) existing.remove();
  ss.setActiveSheet(cashSheet);
  cashSheet.getRange(2, 1, lastRow - 1, Math.max(cmap.mainCat, cmap.expType) + 1).createFilter();
  const filter = cashSheet.getFilter();
  filter.setColumnFilterCriteria(cmap.month, SpreadsheetApp.newFilterCriteria().whenTextContains(month).build());
  const crit = SpreadsheetApp.newFilterCriteria().whenTextContains(selectedValue).build();
  if (isSmall) filter.setColumnFilterCriteria(cmap.expType, crit);
  if (isMain)  filter.setColumnFilterCriteria(cmap.mainCat, crit);
  cashSheet.setActiveSelection("A3");
}

function clearCashFilter() {
  const cashSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cash");
  if (!cashSheet) return;
  const f = cashSheet.getFilter();
  if (f) f.remove();
  cashSheet.activate();
  cashSheet.setActiveSelection("A3");
}

// ============================================================
// HELPERS
// ============================================================
function firstEmptyRowFE_(sh, colMonth, colSpent, startRow) {
  var ncols = Math.max(colMonth, colSpent);
  var vals = sh.getRange(startRow, 1, sh.getMaxRows() - startRow + 1, ncols).getValues();
  for (var i = 0; i < vals.length; i++) {
    var mv = vals[i][colMonth - 1];
    var sv = vals[i][colSpent - 1];
    if ((mv === "" || mv === null) && (sv === "" || sv === null)) {
      return startRow + i;
    }
  }
  return sh.getMaxRows() + 1;
}

function jsonFE_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getFirstEmptyRow_(sheet, col, startRow) {
  const vals = sheet.getRange(startRow, col, sheet.getMaxRows()-startRow+1, 1).getDisplayValues();
  for (let i=0; i<vals.length; i++) if (vals[i][0]==="") return startRow+i;
  return sheet.getMaxRows()+1;
}

function extractRateFromText_(text) {
  const cleaned = String(text||"").replace(/[٠-٩]/g,d=>"٠١٢٣٤٥٦٧٨٩".indexOf(d))
    .replace(/،/g,".").replace(/,/g,".");
  const m = cleaned.match(/-?\d+(?:\.\d+)?/g);
  if (!m||!m.length) return null;
  const v = Number(m[m.length-1]);
  return isNaN(v)?null:v;
}

function getLinkedRowFromNote_(cell, key) {
  const note = cell.getNote();
  if (!note) return null;
  const m = note.match(new RegExp(key+":(\\d+)"));
  return m ? Number(m[1]) : null;
}

function setLinkedRowNote_(cell, key, rowNumber) {
  const lines = (cell.getNote()||"").split("\n")
    .filter(l=>l.trim()!==""&&!l.startsWith(key+":"));
  lines.push(key+":"+rowNumber);
  cell.setNote(lines.join("\n"));
}

function normalizeMonth_(v) {
  const map = {"1":"January","2":"February","3":"March","4":"April","5":"May","6":"June",
               "7":"July","8":"August","9":"September","10":"October","11":"November","12":"December"};
  return map[String(v).trim()] || String(v).trim();
}

function cleanCategory_(value) {
  return String(value).normalize("NFKC")
    .replace(/[​-‏؜‪-‮⁦-⁩﻿]/g,"")
    .replace(/[^\w\s&\-]/g,"").replace(/\s+/g," ").trim();
}

function num(v) {
  return (typeof v === "number" && !isNaN(v)) ? v : 0;
}

// ============================================================
// TEST HELPERS (run manually from the Apps Script editor)
// ============================================================
function testGemini() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(askGemini_(ss, "صرفت كام الشهر ده وفي إيه أكتر فئة؟"));
}

function testBalances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(JSON.stringify(computeAccountBalances_(ss)));
}

function testWealthSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var thisMonth = Utilities.formatDate(new Date(), tz, "MMMM");
  Logger.log(JSON.stringify({
    balances:    computeAccountBalances_(ss),
    totals:      computeMonthTotals_(ss, thisMonth),
    dollarHold:  computeDollarHoldings_(ss),
    dollarPrice: getDollarPrice_(ss),
    netWorth:    computeNetWorthAssets_(ss),
    loans:       computeLoansTotal_(ss)
  }, null, 2));
}

function testInstaOverrideAndMonthSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledgerInsta = (computeAccountBalances_(ss)["Insta"] || 0);
  var override = getInstaBalanceOverride_(ss);
  var dashSh = ss.getSheetByName("Dashboard");
  var dashMonth = dashSh ? String(dashSh.getRange("A7").getValue() || "").trim() : "(no Dashboard sheet)";
  Logger.log(JSON.stringify({
    ledgerInsta:        ledgerInsta,
    financeAA19Override: override,
    usedValue:          (override !== null ? override : ledgerInsta),
    dashboardA7Month:   dashMonth
  }, null, 2));
}
