/***********************************************************************
 *  FINANCE QUICK-ENTRY  +  BALANCES  +  AI ASK (Gemini)
 *  Kiwan's Finance Tracker
 *  Sheet ID: 1kwogE9AbLV5tQAPkni_HvgMjIFxdTh-B89vVAm1KCF0
 *
 *  STEP: set your Gemini key in Script Properties, NOT here — this repo
 *  is public, so anything pasted into this file is visible to anyone,
 *  forever (git history keeps it even after a later edit removes it).
 *    Apps Script editor → Project Settings → Script properties → Add:
 *      GEMINI_API_KEY = <your real key from aistudio.google.com/apikey>
 *    Optional overrides (all have safe defaults if unset):
 *      GEMINI_MODEL, MONTHLY_LIMIT
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
 *  Apps Script project — a second copy in Code.gs is exactly what used
 *  to cause the API to silently serve inconsistent/incorrect data,
 *  depending on file load order. Code.gs now only holds sheet-side
 *  automation (onEdit triggers, the Dashboard menu) and calls the
 *  helpers defined here (getCashColMap_, getDollarColMap_, etc.) —
 *  every .gs file in a project shares one global scope, so that works
 *  without duplicating anything. Do not add another doGet/doPost, or a
 *  second copy of any function already defined here, anywhere else.
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
 * Labels used to find each account's running balance in the Finance
 * sheet's "ACCOUNT BALANCES" boxes (see getAccountBalanceOverride_).
 * No QNB entry — the Finance sheet has no QNB box, so that account is
 * always the ledger sum (currently 0, since no transactions use it yet).
 */
var ACCOUNT_FINANCE_LABELS = {
  Insta: ["NBE/Insta", "NBE / INSTA", "NBE/ INSTA"],
  HSBC:  ["HSBC"],
  Cash:  ["CASH"],
  VF:    ["VF"],
  Telda: ["Telda"],
  Visa:  ["Visa"]
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
// SHEET COLUMN MAPS — header-detected, not fixed column numbers.
// Shared with Code.gs's onEdit triggers and sync handlers.
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
// whose figures live in scattered, hand-formatted boxes rather than a
// tidy header row. We locate values by label TEXT, not a fixed address,
// so the API survives the sheet being reorganized. The Finance tab
// mixes two layouts for its boxes:
//   - label, then its number a cell or two to the RIGHT on the same
//     row (e.g. "Income This Month | ... | 50,066.24")
//   - several sibling labels sharing one row, each with its number
//     directly BELOW it (e.g. "NBE/Insta | HSBC | Cash" headers, with
//     each account's balance on the next row under its own header;
//     merged cells can also shift the number one column left/right of
//     its header when exported, e.g. the "CASH" box)
// This single helper tries both, so one call handles every box on the
// sheet without needing to know which layout a given label uses.
// ============================================================
function findLabeledNumber_(sheet, labels, maxRows, maxCols) {
  if (!sheet) return null;
  var lr = Math.min(sheet.getLastRow(), maxRows || 80);
  var lc = Math.min(sheet.getLastColumn(), maxCols || 60);
  if (lr < 1 || lc < 1) return null;
  var vals = sheet.getRange(1, 1, lr, lc).getValues();
  var norm = function(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); };
  var wanted = labels.map(norm);

  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var cell = norm(vals[r][c]);
      if (!cell || wanted.indexOf(cell) === -1) continue;

      // 1) same row, scanning right: first numeric cell wins. Stop as
      //    soon as we hit ANY other non-empty cell that isn't numeric
      //    (a sibling box's label) so we never walk into the wrong box.
      for (var c2 = c + 1; c2 < vals[r].length; c2++) {
        var n1 = parseMoneyCell_(vals[r][c2]);
        if (!isNaN(n1)) return n1;
        if (vals[r][c2] !== "" && vals[r][c2] !== null) break;
      }

      // 2) one or two rows below, same column ±1 (absorbs the
      //    occasional merged-cell column shift).
      for (var dr = 1; dr <= 2; dr++) {
        var rr = r + dr;
        if (rr >= vals.length) break;
        for (var dc = -1; dc <= 1; dc++) {
          var cc = c + dc;
          if (cc < 0 || cc >= vals[rr].length) continue;
          var n2 = parseMoneyCell_(vals[rr][cc]);
          if (!isNaN(n2)) return n2;
        }
      }
    }
  }
  return null;
}

/* Parses a cell that may be a live number OR currency-formatted text
 * (e.g. "118,927.88 EGP", "-£210.00") into a plain number. */
function parseMoneyCell_(v) {
  if (typeof v === "number") return v;
  if (v === null || v === undefined || v === "") return NaN;
  var s = String(v).trim();
  var negParen = /^\(.*\)$/.test(s);
  var cleaned = s.replace(/[^0-9.\-]/g, "");
  if (!cleaned) return NaN;
  var n = Number(cleaned);
  if (isNaN(n)) return NaN;
  return negParen ? -Math.abs(n) : n;
}

/* Looks up a figure on the Finance sheet by label (see
 * findLabeledNumber_). Returns null if the Finance sheet or the label
 * isn't found, so callers can fall back to a ledger computation. */
function getFinanceValue_(ss, labels) {
  var fin = ss.getSheetByName("Finance");
  if (!fin) return null;
  return findLabeledNumber_(fin, labels);
}

/* Account balance for one of the "ACCOUNT BALANCES" boxes (NBE/Insta,
 * HSBC, Cash, VF, Telda, Visa), read straight from the Finance sheet
 * instead of the bottom-up Cash-ledger sum. These are running totals
 * "as of now", so this override applies no matter which month the app
 * is currently browsing. Falls back to the ledger sum (already computed
 * in computeAccountBalances_) if the label can't be found. */
function getAccountBalanceOverride_(ss, labels) {
  return getFinanceValue_(ss, labels);
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

/* Average monthly spend. Prefers the Finance sheet's "AVG MONTHLY
 * SPEND" box; falls back to summing the Avg sheet's Q3:Q11 block (a
 * hand-built summary column, not a header-based table, so it can't be
 * header-detected the way Cash/Dollar are). */
function computeAvgSpend_(ss) {
  var v = getFinanceValue_(ss, ["avg monthly spend", "average monthly spending"]);
  if (v !== null) return Math.abs(v);

  var a = ss.getSheetByName("Avg");
  if (a) {
    try {
      var vals = a.getRange("Q3:Q11").getValues();
      var total = 0;
      vals.forEach(function(r) {
        var vv = Math.abs(Number(r[0]) || 0);
        if (vv > 0) total += vv;
      });
      if (total > 0) return total;
    } catch (e) {
      Logger.log("computeAvgSpend_: Avg!Q3:Q11 read failed: " + e);
    }
  }
  return 0;
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

/* Always all 12 calendar months, in order — not just the ones with
 * Cash-sheet entries so far. The month picker should let the user
 * browse ahead/behind regardless of whether that month has ledger data
 * yet (the Avg-sheet breakdown and ledger computations both handle an
 * empty month gracefully, showing "No data" rather than erroring). */
function getAvailableMonths_(ss) {
  return ["January","February","March","April","May","June","July",
          "August","September","October","November","December"];
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

/* EGP-per-USD rate. Prefers the Finance sheet's "ACCOUNT BALANCES ▸
 * DOLLAR HOLDINGS" box, where the rate sits directly under the "USDT"
 * unit tag (no dedicated text label exists for it in that box, so we
 * search for "USDT" itself). No hardcoded fallback number — a stale
 * hardcoded rate is exactly the kind of "hardcoded financial value"
 * this project was full of. Falls back to the Net worth sheet, then to
 * 0 (logged) rather than guessing. */
function getDollarPrice_(ss) {
  var v = getFinanceValue_(ss, ["usdt", "dollar price", "dollar"]);
  if (v !== null && v > 0) return v;
  var nw = ss.getSheetByName("Net worth");
  var v2 = findLabeledNumber_(nw, ["dollar price", "live dollar", "live dollar price"]);
  if (v2 !== null && v2 > 0) return v2;
  Logger.log("getDollarPrice_: no dollar price found on Finance or Net worth sheets");
  return 0;
}

/* Net worth + assets. Prefers the Finance sheet's "NET WORTH" / "Assets"
 * boxes (validated against the user's Finance-tab export); falls back
 * to the "Net worth" sheet's labeled rows ("Total Wealth" / "Assets"). */
function computeNetWorthAssets_(ss) {
  var netWorth = getFinanceValue_(ss, ["net worth"]);
  var assetsVal = getFinanceValue_(ss, ["assets", "assets value", "assets:"]);

  if (netWorth === null || assetsVal === null) {
    var nw = ss.getSheetByName("Net worth");
    if (nw) {
      var lr = nw.getLastRow();
      if (lr > 0) {
        var vals = nw.getRange(1, 1, lr, 2).getValues();
        vals.forEach(function(r) {
          var label = String(r[0] || "").trim().toLowerCase();
          if (netWorth === null && label === "total wealth") netWorth = Number(r[1]) || 0;
          if (assetsVal === null && label === "assets")       assetsVal = Number(r[1]) || 0;
        });
      }
    }
  }
  return { netWorth: netWorth || 0, assetsVal: assetsVal || 0 };
}

/* Net loans/debt figure. Prefers the Finance sheet's "Loans/Debt"
 * summary cell; falls back to the Loans sheet's own Grand Total row
 * (EGP column) if the Finance sheet doesn't have it. Negative = money
 * owed TO Kiwan; positive = Kiwan owes it. */
function computeLoansTotal_(ss) {
  var v = getFinanceValue_(ss, ["loans/debt", "loans / debt", "loans"]);
  if (v !== null) return v;

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

/* Income/expense "In"/"Out" breakdown sourced from the "Avg" sheet,
 * which has one column per calendar month (January..December) plus a
 * trailing "Avg" column — unlike the Main Categories/Finance sheets,
 * which only ever describe whichever month is currently "live", this
 * lets the app show a category breakdown for ANY month, driven by the
 * app's existing month picker (no new UI needed).
 *
 * Row layout matches Main Categories: "Income" followed by its flat
 * "* Detail" rows, then "Expenses" as main categories (Bills, Food,
 * ...) each followed by their own "* Detail" rows, ending at the
 * "sum" row (before the sheet's second, chart-data table below it).
 * The sheet spells April as "Abril", handled via MONTH_ALIASES.
 *
 * KNOWN DATA-QUALITY CAVEAT (sheet-side, not fixable from here): this
 * sheet's Income column appears to require an exact category-text
 * match, so it misses at least one real spelling variant used in the
 * Cash sheet ("Parents Support" vs "Parental Support") — e.g. its July
 * total undercounts by exactly that amount vs the Finance/Main
 * Categories tabs. Expenses aren't affected. Mirrored faithfully as
 * requested; computeMonthBreakdown_ (bottom-up from the ledger) is the
 * accurate alternative if that matters more than matching this sheet.
 *
 * Returns null if the sheet, the month's column, or any real data for
 * it can't be found, so callers can fall back to the ledger breakdown.
 */
function getAvgSheetBreakdown_(ss, monthName) {
  var sh = ss.getSheetByName("Avg");
  if (!sh) return null;
  var lr = sh.getLastRow();
  if (lr < 3) return null;
  var lc = sh.getLastColumn();

  var MONTH_ALIASES = { "abril": "april", "noc": "november" };
  var norm = function(s) {
    s = String(s || "").trim().toLowerCase();
    return MONTH_ALIASES[s] || s;
  };
  var wantMonth = norm(monthName);

  var monthCol = -1;
  var headerScan = sh.getRange(1, 1, Math.min(3, lr), lc).getValues();
  headerFind:
  for (var hr = 0; hr < headerScan.length; hr++) {
    for (var hc = 0; hc < headerScan[hr].length; hc++) {
      if (norm(headerScan[hr][hc]) === wantMonth) { monthCol = hc + 1; break headerFind; }
    }
  }
  if (monthCol === -1) return null;

  var vals = sh.getRange(1, 1, lr, monthCol).getValues();
  var income = {}, expense = {};
  var section = null;
  var currentMainCat = null;
  var foundAnyItem = false;

  for (var r = 0; r < vals.length; r++) {
    var rawName = String(vals[r][0] || "").trim();
    if (!rawName) continue;
    var nameLower0 = rawName.toLowerCase();
    if (nameLower0 === "sum") break; // stop before the second (chart) table

    var amt = parseMoneyCell_(vals[r][monthCol - 1]);
    if (isNaN(amt)) amt = 0;

    var isDetail = rawName.indexOf("*") === 0;
    var name = isDetail ? rawName.replace(/^\*+\s*/, "").trim() : rawName;
    var nameL = name.toLowerCase();

    if (!isDetail) {
      if (nameL === "income") {
        section = "income";
        currentMainCat = "Income";
        income[currentMainCat] = { total: Math.abs(amt), items: {} };
        continue;
      }
      if (nameL === "expenses" || nameL === "expense") {
        section = "expense";
        currentMainCat = null;
        continue;
      }
      if (section === "expense") {
        currentMainCat = name;
        expense[currentMainCat] = { total: Math.abs(amt), items: {} };
      }
      continue;
    }

    if (amt === 0 || !currentMainCat) continue;
    foundAnyItem = true;
    if (section === "income")  income[currentMainCat].items[name]  = Math.abs(amt);
    if (section === "expense") expense[currentMainCat].items[name] = Math.abs(amt);
  }

  // No real transactions found for this month's column (e.g. a future
  // month the sheet hasn't populated yet) -> let the caller fall back
  // to the ledger instead of showing a misleadingly empty breakdown.
  if (!foundAnyItem) return null;

  return { income: income, expense: expense };
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

    // Account balances: bottom-up ledger sum first, then override with
    // the Finance sheet's own "ACCOUNT BALANCES" boxes where available
    // (these are running totals "as of now", so this applies regardless
    // of which month is being browsed).
    var balances = computeAccountBalances_(ss);
    Object.keys(ACCOUNT_FINANCE_LABELS).forEach(function(key) {
      var override = getAccountBalanceOverride_(ss, ACCOUNT_FINANCE_LABELS[key]);
      if (override !== null) {
        Logger.log("handleGet_: overriding " + key + " balance with Finance-sheet value " +
          override + " (ledger sum was " + (balances[key] || 0) + ")");
        balances[key] = override;
      }
    });

    var totals    = computeMonthTotals_(ss, thisMonth);
    var breakdown = computeMonthBreakdown_(ss, thisMonth);
    var avgSpend  = computeAvgSpend_(ss);
    var allMonths = getAvailableMonths_(ss);

    // The "Avg" sheet has one column per calendar month, so the In/Out
    // breakdown can be sourced from it for whichever month is being
    // browsed, not just the live one — falls back to the ledger-computed
    // breakdown above if that month's column isn't found/populated.
    var avgBreakdown = getAvgSheetBreakdown_(ss, thisMonth);
    if (avgBreakdown) {
      Logger.log("handleGet_: sourcing income/expense breakdown from Avg sheet for " + thisMonth);
      breakdown = avgBreakdown;
    }

    // The Finance sheet's "THIS MONTH" box (Income This Month / Total
    // Expenses / Saved) only describes the live current month, so it
    // only overrides the ledger totals when that's the month being
    // shown — browsing to a past month always uses the ledger sum for
    // that month instead.
    var realCurrentMonth = Utilities.formatDate(new Date(), tz, "MMMM");
    if (thisMonth.toLowerCase() === realCurrentMonth.toLowerCase()) {
      var incomeOverride  = getFinanceValue_(ss, ["income this month"]);
      var expenseOverride = getFinanceValue_(ss, ["total expenses"]);
      var savedOverride   = getFinanceValue_(ss, ["saved / (spent)", "saved/(spent)", "saved"]);
      if (incomeOverride !== null || expenseOverride !== null || savedOverride !== null) {
        Logger.log("handleGet_: overriding this-month totals with Finance-sheet values " +
          "(income=" + incomeOverride + " expense=" + expenseOverride + " saved=" + savedOverride +
          "), ledger computed income=" + totals.income + " expense=" + totals.expense);
      }
      if (incomeOverride  !== null) totals.income  = incomeOverride;
      if (expenseOverride !== null) totals.expense = expenseOverride;
      totals.saved = (savedOverride !== null) ? savedOverride : (totals.income - totals.expense);
    }

    var dollarHoldingsOverride = getFinanceValue_(ss, ["dollar holdings"]);
    var dollarHoldings = (dollarHoldingsOverride !== null) ? dollarHoldingsOverride : computeDollarHoldings_(ss);
    var dollarPrice    = getDollarPrice_(ss);
    var netWorthAssets = computeNetWorthAssets_(ss);
    var loans          = computeLoansTotal_(ss);

    var accountsSum = 0;
    Object.keys(balances).forEach(function(k) { accountsSum += balances[k]; });

    var availableBalance = getFinanceValue_(ss, ["available balance"]);
    if (availableBalance === null) {
      availableBalance = accountsSum + dollarHoldings * dollarPrice;
    }

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
// HELPERS (doGet/doPost only — Code.gs has its own set for the
// onEdit sync handlers)
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

function testFinanceOverrides() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var thisMonth = Utilities.formatDate(new Date(), tz, "MMMM");

  var ledgerBalances = computeAccountBalances_(ss);
  var accountOverrides = {};
  Object.keys(ACCOUNT_FINANCE_LABELS).forEach(function(key) {
    accountOverrides[key] = {
      ledger:   ledgerBalances[key] || 0,
      override: getAccountBalanceOverride_(ss, ACCOUNT_FINANCE_LABELS[key])
    };
  });

  var dashSh = ss.getSheetByName("Dashboard");
  var dashMonth = dashSh ? String(dashSh.getRange("A7").getValue() || "").trim() : "(no Dashboard sheet)";

  Logger.log(JSON.stringify({
    accountOverrides:  accountOverrides,
    thisMonthLedger:   computeMonthTotals_(ss, thisMonth),
    thisMonthOverride: {
      income:  getFinanceValue_(ss, ["income this month"]),
      expense: getFinanceValue_(ss, ["total expenses"]),
      saved:   getFinanceValue_(ss, ["saved / (spent)", "saved/(spent)", "saved"])
    },
    availableBalanceOverride: getFinanceValue_(ss, ["available balance"]),
    dollarHoldingsOverride:   getFinanceValue_(ss, ["dollar holdings"]),
    dollarPrice:              getDollarPrice_(ss),
    netWorthAssets:           computeNetWorthAssets_(ss),
    loans:                    computeLoansTotal_(ss),
    avgSpend:                 computeAvgSpend_(ss),
    dashboardA7Month:         dashMonth
  }, null, 2));
}

function testAvgSheetBreakdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var months = ["January","February","March","April","May","June","July",
                "August","September","October","November","December"];
  var out = {};
  months.forEach(function(m) {
    out[m] = getAvgSheetBreakdown_(ss, m);
  });
  Logger.log(JSON.stringify(out, null, 2));
}
