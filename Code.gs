/***********************************************************************
 *  SHEET-SIDE AUTOMATION — Kiwan's Finance Tracker
 *
 *  Runs inside the spreadsheet itself: the onEdit trigger that
 *  auto-fills month/date and keeps the Cash/Dollar sheets in sync when
 *  a "Dollar Deposit"/"Transfer Out" row is entered, plus the Dashboard
 *  Tools menu.
 *
 *  This file has NO doGet/doPost — the web app entry point lives
 *  entirely in FinanceEntry.gs. Do not add one here; two files each
 *  defining doGet/doPost is exactly what used to make the API silently
 *  serve inconsistent data depending on file load order.
 *
 *  getCashColMap_ / getDollarColMap_ (and everything else this file
 *  calls that isn't defined below) live in FinanceEntry.gs — every .gs
 *  file in an Apps Script project shares one global scope, so no import
 *  is needed, but don't redefine any of those names here.
 ***********************************************************************/

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
// AUTO DATE/MONTH + SYNC — Cash, Dollar, AnQa, Rawaq
// ============================================================
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
// HELPERS (onEdit sync handlers only — FinanceEntry.gs has its own
// set for doGet/doPost)
// ============================================================
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
