// ============================================================
// SETUP (run once to create the onEdit trigger if needed)
// ============================================================
function setup() {
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// ============================================================
// AUTO DATE/MONTH + SYNC — Cash, Dollar
// ============================================================
function onEdit(e) {
  if (!e || !e.range) return;
  const sh  = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const val = String(e.range.getDisplayValue()).trim();

  // Detect actual column indexes dynamically for Cash and Dollar
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
    try {
      return getCashColMap_(sheet).spent;
    } catch(e) { return 3; }
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

// ============================================================
// DASHBOARD MENU TOOLS
// ============================================================
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
// DASHBOARD API — supports ?month=May&year=2026
// ============================================================
function doGet(e) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cash = ss.getSheetByName("Cash");
  const avg  = ss.getSheetByName("Avg");
  const nw   = ss.getSheetByName("Net worth");

  const reqMonth = (e && e.parameter && e.parameter.month) ? e.parameter.month : null;
  const reqYear  = (e && e.parameter && e.parameter.year)  ? e.parameter.year  : null;

  var serveMonth, serveYear;
  if (reqMonth && reqYear) {
    serveMonth = reqMonth;
    serveYear  = reqYear;
  } else {
    const tz = ss.getSpreadsheetTimeZone();
    serveMonth = Utilities.formatDate(new Date(), tz, "MMMM");
    serveYear  = Utilities.formatDate(new Date(), tz, "yyyy");
  }

  const cmap    = getCashColMap_(cash);
  const lastRow = cash.getLastRow();
  if (lastRow < 3) {
    var cb0 = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "dashboardCallback";
    return ContentService.createTextOutput(cb0 + "({})").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  const totalCols = Math.max(cmap.month, cmap.date, cmap.spent, cmap.account,
                              cmap.description, cmap.expCat, cmap.expType, cmap.mainCat);
  const cashBlock = cash.getRange(3, 1, lastRow - 2, totalCols).getValues();

  const iMonth  = cmap.month       - 1;
  const iDate   = cmap.date        - 1;
  const iSpent  = cmap.spent       - 1;
  const iAcc    = cmap.account     - 1;
  const iDesc   = cmap.description - 1;
  const iExpCat = cmap.expCat      - 1;
  const iExpTyp = cmap.expType     - 1;
  const iMainCt = cmap.mainCat     - 1;

  const INCOME_TYPES = new Set([
    "income", "paycheck", "indriver", "side income", "bonus",
    "parental support", "salary"
  ]);
  const SKIP_CATS = new Set([
    "money movement", "deposit", "payback", "refund",
    "insta withdrawal", "dollar withdrawal", "dollar deposit",
    "transfer=>insta", "transfer=>vf", "transfer=>telda",
    "transfer=>cash", "transfer=>qnb"
  ]);

  const monthRows = cashBlock.filter(function(r) {
    return String(r[iMonth] || "").toLowerCase() === serveMonth.toLowerCase();
  });

  var totalIncome = 0, totalExpense = 0;
  var catTotals   = {};
  var typeTotals  = {};
  var transactions = [];
  var tz = Session.getScriptTimeZone();

  monthRows.forEach(function(r) {
    var spent   = Number(r[iSpent]) || 0;
    var expCat  = String(r[iExpCat] || "").trim();
    var expType = String(r[iExpTyp] || "").trim();
    var mainCat = String(r[iMainCt] || "").trim();
    var desc    = String(r[iDesc] || "-");
    var dateVal = r[iDate];
    var expCatL = expCat.toLowerCase();
    var expTypL = expType.toLowerCase();

    if (SKIP_CATS.has(expCatL) || SKIP_CATS.has(expTypL)) return;

    if (INCOME_TYPES.has(expCatL) || INCOME_TYPES.has(expTypL) ||
        INCOME_TYPES.has(mainCat.toLowerCase())) {
      if (spent > 0) totalIncome += spent;
    } else if (spent < 0) {
      totalExpense += Math.abs(spent);
      if (mainCat) catTotals[mainCat] = (catTotals[mainCat] || 0) + Math.abs(spent);
      if (expType) typeTotals[expType] = (typeTotals[expType] || 0) + Math.abs(spent);
    }

    if (expCat || mainCat) {
      var dateStr = "";
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, tz, "dd MMM");
      } else {
        dateStr = String(dateVal || "").substring(0, 8);
      }
      transactions.push({
        date: dateStr, desc: desc.substring(0, 40),
        amount: spent, cat: expType || expCat, main: mainCat || expCat
      });
    }
  });

  transactions.reverse();
  transactions = transactions.slice(0, 20);

  var catOrder1 = ["Food","Car","Health","Sports","Personal","Bills","Finance","Others","Drinks"];
  var expenses = catOrder1
    .filter(function(k) { return catTotals[k] && catTotals[k] > 0; })
    .map(function(k) {
      return { name: k, actual: catTotals[k], planned: 0, diff: -catTotals[k] };
    }).sort(function(a,b) { return b.actual - a.actual; });

  var catOrder2 = ["Income","Bills","Food","Drinks","Sports","Car","Health","Personal","Finance","Others"];
  var avgBlock = [];
  try { avgBlock = avg.getRange("Q2:Q11").getValues(); } catch(ex) { avgBlock = catOrder2.map(function(){ return [0]; }); }
  var main_cats = catOrder2.map(function(name, i) {
    var actual = name === "Income" ? totalIncome : -(catTotals[name] || 0);
    return { name: name, avg: Number((avgBlock[i]||[0])[0]) || 0, actual: actual };
  });

  var incomeTypes = {};
  monthRows.forEach(function(r) {
    var expCat  = String(r[iExpCat] || "").trim();
    var expType = String(r[iExpTyp] || "").trim();
    var spent   = Number(r[iSpent]) || 0;
    var mainCat = String(r[iMainCt] || "").trim();
    if ((INCOME_TYPES.has(expCat.toLowerCase()) || INCOME_TYPES.has(expType.toLowerCase()) ||
         INCOME_TYPES.has(mainCat.toLowerCase())) && spent > 0) {
      var label = expType || expCat || mainCat;
      incomeTypes[label] = (incomeTypes[label] || 0) + spent;
    }
  });
  var incomes = Object.keys(incomeTypes).map(function(k) {
    return { name: k, actual: incomeTypes[k], planned: 0 };
  }).sort(function(a,b) { return b.actual - a.actual; });

  var netWorth = 0, assetsVal = 0;
  if (nw) {
    var nwLr = nw.getLastRow();
    if (nwLr > 0) {
      var nwVals = nw.getRange(1, 1, nwLr, 2).getValues();
      nwVals.forEach(function(r) {
        var label = String(r[0] || "").trim().toLowerCase();
        if (label === "total wealth") netWorth = Number(r[1]) || 0;
        if (label === "assets")       assetsVal = Number(r[1]) || 0;
      });
    }
  }

  var dollarSh = ss.getSheetByName("Dollar");
  var dollarHoldings = 0;
  var dollarPrice    = 51.35;
  if (dollarSh) {
    var dcol = getDollarColMap_(dollarSh);
    var dlr  = dollarSh.getLastRow();
    if (dlr > 2) {
      var dBlock = dollarSh.getRange(3, dcol.spent, dlr - 2, 1).getValues();
      dBlock.forEach(function(r){ dollarHoldings += Number(r[0]) || 0; });
    }
  }

  if (nw) {
    var nwLr2 = nw.getLastRow();
    var nwVals2 = nw.getRange(1, 1, nwLr2, 2).getValues();
    nwVals2.forEach(function(r) {
      var label = String(r[0] || "").trim().toLowerCase();
      if (label.indexOf("live dollar") !== -1) {
        var v = Number(nw.getRange(r.indexOf(r), 3).getValue());
        if (!v) v = Number(r[1]) || 0;
        if (v > 0) dollarPrice = v;
      }
    });
    try {
      var p = Number(nw.getRange("C2").getValue());
      if (p > 0) dollarPrice = p;
    } catch(ex){}
  }

  var accountTotals = {};
  cashBlock.forEach(function(r) {
    var rawAcc = String(r[iAcc] || "").trim();
    if (!rawAcc) return;
    var amt = Number(r[iSpent]) || 0;
    if (amt === 0) return;
    accountTotals[rawAcc] = (accountTotals[rawAcc] || 0) + amt;
  });

  var ALIASES = {
    "nbe": "Insta", "insta": "Insta", "nbe/insta": "Insta",
    "hsbc": "HSBC", "cash": "Cash", "vf": "VF", "vf visa": "VF",
    "telda": "Telda", "qnb": "Qnb", "visa": "Visa"
  };
  var namedBalances = {};
  Object.keys(accountTotals).forEach(function(raw) {
    var key = ALIASES[raw.toLowerCase()] || raw;
    namedBalances[key] = (namedBalances[key] || 0) + accountTotals[raw];
  });

  var mList   = ["November","December","January","February","March","April","May","June"];
  var mLabels = ["Nov 25","Dec 25","Jan 26","Feb 26","Mar 26","Apr 26","May 26","Jun 26"];
  var months_trend = mList.map(function(mn, i) {
    var inc = 0, exp = 0;
    cashBlock.forEach(function(r) {
      var rm  = String(r[iMonth] || "").toLowerCase();
      var ec  = String(r[iExpCat] || "").trim().toLowerCase();
      var et  = String(r[iExpTyp] || "").trim().toLowerCase();
      var mc  = String(r[iMainCt] || "").trim().toLowerCase();
      var sp  = Number(r[iSpent]) || 0;
      if (rm !== mn.toLowerCase()) return;
      if (SKIP_CATS.has(ec) || SKIP_CATS.has(et)) return;
      if (INCOME_TYPES.has(ec) || INCOME_TYPES.has(et) || INCOME_TYPES.has(mc)) { if (sp > 0) inc += sp; }
      else if (sp < 0) exp += Math.abs(sp);
    });
    return { month: mLabels[i], income: inc, expense: exp, surplus: inc - exp };
  }).filter(function(m){ return m.income > 0 || m.expense > 0; });

  var allMonths = ["November","December","January","February","March","April","May","June",
                   "July","August","September","October"];
  var mIdx = allMonths.map(function(m){ return m.toLowerCase(); }).indexOf(serveMonth.toLowerCase());
  var runningBalance = 0;
  cashBlock.forEach(function(r) {
    var rm  = String(r[iMonth] || "").toLowerCase();
    var ec  = String(r[iExpCat] || "").trim().toLowerCase();
    var et  = String(r[iExpTyp] || "").trim().toLowerCase();
    var sp  = Number(r[iSpent]) || 0;
    var rIdx = allMonths.map(function(m){ return m.toLowerCase(); }).indexOf(rm);
    if (SKIP_CATS.has(ec) || SKIP_CATS.has(et)) return;
    if (rIdx <= mIdx) runningBalance += sp;
  });

  var seen = {}, uniqueMonths = [];
  cashBlock.forEach(function(r) {
    var mn = String(r[iMonth] || "").trim();
    if (mn && !seen[mn]) { seen[mn] = true; uniqueMonths.push(mn); }
  });

  var avgExpense = 0;
  try {
    var avgVals = avg.getRange("Q3:Q11").getValues();
    avgVals.forEach(function(r){ avgExpense += Math.abs(Number(r[0]) || 0); });
  } catch(ex){}

  var data = {
    month:             serveMonth,
    year:              serveYear,
    available_months:  uniqueMonths,
    available_balance: dollarHoldings * dollarPrice + (namedBalances["Insta"] || 0) +
                       (namedBalances["Cash"] || 0) + (namedBalances["HSBC"] || 0) +
                       (namedBalances["VF"] || 0) + (namedBalances["Telda"] || 0),
    dollar_holdings:   dollarHoldings,
    dollar_price:      dollarPrice,
    loans:             0,
    avg_spend:         avgExpense / 9 || 0,
    qnb_insta:  namedBalances["Insta"]  || 0,
    cash_phys:  namedBalances["Cash"]   || 0,
    vf_visa:    namedBalances["VF"]     || 0,
    visa_misr:  namedBalances["Telda"]  || 0,
    hsbc:       namedBalances["HSBC"]   || 0,
    due_to_anqa: 0,
    net_worth:   netWorth,
    assets_val:  assetsVal,
    start_bal:   runningBalance,
    end_bal:     runningBalance,
    saved:       totalIncome - totalExpense,
    pct_change:  totalIncome > 0 ? (totalIncome - totalExpense) / totalIncome : 0,
    total_exp_planned: 0,
    total_exp_actual:  totalExpense,
    total_inc_planned: 0,
    total_inc_actual:  totalIncome,
    expenses:    expenses,
    incomes:     incomes,
    main_cats:   main_cats,
    months_trend: months_trend,
    transactions: transactions
  };

  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "dashboardCallback";
  return ContentService
    .createTextOutput(cb + "(" + JSON.stringify(data) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ============================================================
// SHARED COLUMN MAP HELPERS
// ============================================================
function getCashColMap_(sh) {
  var headers = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = { month:1, date:2, spent:3, account:4, description:5, priority:6, expCat:7, expType:8, mainCat:9, notes:10 };
  var nameMap = {
    "month":"month","date":"date","spent":"spent","account":"account",
    "description":"description","priority":"priority",
    "expenses category":"expCat","expense type":"expType","main category":"mainCat","notes":"notes"
  };
  headers.forEach(function(h,i){
    var key=nameMap[String(h||"").trim().toLowerCase()];
    if(key) map[key]=i+1;
  });
  return map;
}

function getDollarColMap_(sh) {
  var headers = sh.getRange(2, 1, 1, Math.min(sh.getLastColumn(), 10)).getValues()[0];
  var map = { month:1, date:2, spent:3, description:4, dollarType:5 };
  var nameMap = {
    "month":"month","date":"date","spent":"spent","description":"description",
    "transfer/profit":"dollarType","type":"dollarType"
  };
  headers.forEach(function(h,i){
    var key=nameMap[String(h||"").trim().toLowerCase()];
    if(key) map[key]=i+1;
  });
  return map;
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
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

function num(v) {
  return (typeof v === "number" && !isNaN(v)) ? v : 0;
}
