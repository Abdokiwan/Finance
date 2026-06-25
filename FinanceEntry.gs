/***********************************************************************
 *  FINANCE QUICK-ENTRY  +  BALANCES  +  AI ASK (Gemini)
 *  Kiwan's Finance Tracker
 *  Sheet ID: 1kwogE9AbLV5tQAPkni_HvgMjIFxdTh-B89vVAm1KCF0
 *  STEP: put your Gemini key between the quotes below
 ***********************************************************************/
var GEMINI_API_KEY = "PUT_YOUR_GEMINI_KEY_HERE";  // Set this in Apps Script, never commit the real key
var GEMINI_MODEL   = "gemini-2.5-flash";

// ====== Monthly budget limit (EGP) ======
var MONTHLY_LIMIT = 15000;

/*
 * ACCOUNT NAME MAP — maps every variant found in the Cash sheet → display key.
 * Add new aliases here if account names change in the sheet.
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
 * Categories skipped when computing income/expense totals (internal transfers).
 * Kept in sync with Code.gs SKIP_CATS.
 */
var SKIP_CATS_SET = {
  "money movement":1, "deposit":1, "payback":1, "refund":1,
  "insta withdrawal":1, "dollar withdrawal":1, "dollar deposit":1,
  "transfer=>insta":1, "transfer=>vf":1, "transfer=>telda":1,
  "transfer=>cash":1, "transfer=>qnb":1
};

/*
 * Income type keywords (lower-case). Kept in sync with Code.gs INCOME_TYPES.
 */
var INCOME_KEYS = {
  "income":1, "paycheck":1, "indriver":1, "side income":1,
  "bonus":1, "parental support":1, "salary":1
};

/* ---------- CASH SHEET COLUMN MAP (header-based, auto-detected) ---------- */
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

/* ---------- DOLLAR SHEET COLUMN MAP ---------- */
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

/* ---------- Compute account balances from Cash sheet ----------
 * Sums ALL rows across all months. Uses ACCOUNT_ALIASES to merge
 * variant account names (nbe/insta → Insta, etc.).
 */
function computeAccountBalances_(ss) {
  var sh = ss.getSheetByName("Cash");
  if (!sh) return {};
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

/* ---------- Compute this-month income / expense / saved from Cash ---------- */
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
    // Skip internal transfers and money movements
    if (SKIP_CATS_SET[expCat] || SKIP_CATS_SET[expType] ||
        expCat.indexOf("transfer=>") !== -1 || expType.indexOf("transfer=>") !== -1) return;
    if (INCOME_KEYS[expCat] || INCOME_KEYS[expType] || INCOME_KEYS[mainCat]) {
      if (amt > 0) income += amt;
    } else if (amt < 0) {
      expense += Math.abs(amt);
    }
  });
  return { income:income, expense:expense, saved:income - expense };
}

/* ---------- Compute spent today from Cash sheet ---------- */
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
    if (ds === todayStr && amt < 0 &&
        !SKIP_CATS_SET[expCat] && !SKIP_CATS_SET[expType] &&
        expCat.indexOf("transfer=>") === -1 && expType.indexOf("transfer=>") === -1) {
      spent += Math.abs(amt);
    }
  });
  return spent;
}

/* ---------- Average monthly spend from Avg sheet ---------- */
function computeAvgSpend_(ss) {
  var a = ss.getSheetByName("Avg");
  if (!a) return 0;
  try {
    var vals = a.getRange("Q3:Q11").getValues();
    var total = 0;
    vals.forEach(function(r) {
      var v = Math.abs(Number(r[0]) || 0);
      if (v > 0) total += v;
    });
    return total;
  } catch(e) { return 0; }
}

/* ============================================================
   doGet: balances + AI ask (JSONP)
   ============================================================ */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "cb";
  var out;

  if (e && e.parameter && e.parameter.action === "ask") {
    out = { status:"ok", answer: askGemini_(ss, e.parameter.q || "") };

  } else if (e && e.parameter && e.parameter.action === "recent") {
    out = recentEntries_(ss);

  } else if (e && e.parameter && e.parameter.action === "voice") {
    out = voiceEntry_(ss, e.parameter.text || "");

  } else if (e && e.parameter && e.parameter.action === "memory") {
    out = buildMemory_(ss);

  } else if (e && e.parameter && e.parameter.action === "update") {
    out = updateRow_(ss, e.parameter);

  } else if (e && e.parameter && e.parameter.action === "delete") {
    out = deleteRow_(ss, e.parameter.sheet, Number(e.parameter.row));

  } else {
    // Default: return balances + month summary
    var tz = ss.getSpreadsheetTimeZone();
    var thisMonth = Utilities.formatDate(new Date(), tz, "MMMM");

    var balances = computeAccountBalances_(ss);
    var totals   = computeMonthTotals_(ss, thisMonth);
    var avgSpend = computeAvgSpend_(ss);

    out = {
      // Account balances — keep old keys so the HTML needs no change
      insta:    balances["Insta"]  || 0,
      cash:     balances["Cash"]   || 0,
      vf:       balances["VF"]     || 0,
      visamisr: balances["Telda"]  || 0,
      hsbc:     balances["HSBC"]   || 0,
      qnb:      balances["Qnb"]    || 0,
      visa:     balances["Visa"]   || 0,
      // Monthly summary
      income:   totals.income,
      expense:  totals.expense,
      saved:    totals.saved,
      // Spend tracking
      avgSpend:    MONTHLY_LIMIT,
      avgSpending: avgSpend,
      spentToday:  computeSpentToday_(ss)
    };
  }

  return ContentService
    .createTextOutput(cb + "(" + JSON.stringify(out) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ============================================================
   doPost: add entry
   ============================================================ */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var tz = ss.getSpreadsheetTimeZone();
    var monthName = Utilities.formatDate(now, tz, "MMMM");

    if (data.sheet === "Cash") {
      var sh = ss.getSheetByName("Cash");
      var col = getCashColMap_(sh);
      var row = firstEmptyRowFE_(sh, col.month, col.spent, 3);
      sh.getRange(row, col.month).setValue(monthName);
      sh.getRange(row, col.date).setValue(now);
      sh.getRange(row, col.spent).setValue(data.amount);
      sh.getRange(row, col.account).setValue(data.account);
      sh.getRange(row, col.description).setValue(data.description);
      sh.getRange(row, col.expType).setValue(data.expenseType);

    } else if (data.sheet === "Dollar") {
      var sh = ss.getSheetByName("Dollar");
      var dcol = getDollarColMap_(sh);
      var row = firstEmptyRowFE_(sh, dcol.month, dcol.spent, 3);
      sh.getRange(row, dcol.month).setValue(monthName);
      sh.getRange(row, dcol.date).setValue(now);
      sh.getRange(row, dcol.spent).setValue(data.amount);
      sh.getRange(row, dcol.description).setValue(data.description);
      sh.getRange(row, dcol.dollarType).setValue(data.dollarType);
    }
    return jsonFE_({ status:"ok" });

  } catch (err) {
    return jsonFE_({ status:"error", message: String(err) });
  }
}

/* ============================================================
   Gemini: AI Ask
   ============================================================ */
function askGemini_(ss, question) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "PUT_YOUR_GEMINI_KEY_HERE") {
    return "Add your Gemini key in the script first.";
  }
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
        if (!name || name === "") return;
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
            GEMINI_MODEL + ":generateContent?key=" + GEMINI_API_KEY;
  var payload = { contents: [{ parts: [{ text: prompt }] }] };
  var res = UrlFetchApp.fetch(url, {
    method:"post", contentType:"application/json",
    payload: JSON.stringify(payload), muteHttpExceptions:true
  });
  var json = JSON.parse(res.getContentText());
  if (json.candidates && json.candidates[0]) {
    return json.candidates[0].content.parts[0].text;
  }
  return "Error: " + (json.error ? json.error.message : "No answer.");
}

/* ============================================================
   VOICE ENTRY (via Gemini)
   ============================================================ */
function voiceEntry_(ss, text) {
  if (!text) return { status:"error", message:"No speech." };
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "PUT_YOUR_GEMINI_KEY_HERE")
    return { status:"error", message:"Gemini key not set." };

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

  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL + ":generateContent?key=" + GEMINI_API_KEY;
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
}

/* ============================================================
   SMART MEMORY: learn description -> type/account
   ============================================================ */
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

/* ============================================================
   RECENT ENTRIES (last 60 days)
   ============================================================ */
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
    // Try d-Mon-yyyy
    var mMatch = datePart.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (mMatch) {
      var months = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
      var mi = months.indexOf(mMatch[2].toLowerCase());
      if (mi >= 0) {
        var dn = Math.floor(Date.UTC(+mMatch[3], mi, +mMatch[1]) / 86400000);
        return { ymd: dn, ts: Date.UTC(+mMatch[3], mi, +mMatch[1]) };
      }
    }
    // Try d/M/yyyy or d-M-yyyy
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

/* ============================================================
   UPDATE ROW
   ============================================================ */
function updateRow_(ss, p) {
  try {
    var name = p.sheet;
    var row = Number(p.row);
    var sh = ss.getSheetByName(name);
    if (!sh || !row) return { status:"error", message:"bad target" };

    if (name === "Dollar") {
      var col = getDollarColMap_(sh);
      if (p.amount      !== undefined) sh.getRange(row, col.spent).setValue(Number(p.amount));
      if (p.description !== undefined) sh.getRange(row, col.description).setValue(p.description);
      if (p.dollarType  !== undefined) sh.getRange(row, col.dollarType).setValue(p.dollarType);
    } else {
      var col = getCashColMap_(sh);
      if (p.amount      !== undefined) sh.getRange(row, col.spent).setValue(Number(p.amount));
      if (p.account     !== undefined) sh.getRange(row, col.account).setValue(p.account);
      if (p.description !== undefined) sh.getRange(row, col.description).setValue(p.description);
      if (p.expenseType !== undefined) sh.getRange(row, col.expType).setValue(p.expenseType);
    }
    return { status:"ok" };
  } catch (err) {
    return { status:"error", message: String(err) };
  }
}

/* ============================================================
   DELETE ROW
   ============================================================ */
function deleteRow_(ss, name, row) {
  try {
    var sh = ss.getSheetByName(name);
    if (!sh || !row) return { status:"error", message:"bad target" };
    sh.deleteRow(row);
    return { status:"ok" };
  } catch (err) {
    return { status:"error", message: String(err) };
  }
}

/* ============================================================
   HELPERS
   ============================================================ */

/*
 * Find first empty row in the sheet where both the month column
 * AND the spent column are blank. Uses dynamic column positions
 * so it works regardless of column layout changes.
 */
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

/* ---------- test helpers ---------- */
function testGemini() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(askGemini_(ss, "صرفت كام الشهر ده وفي إيه أكتر فئة؟"));
}

function testBalances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(JSON.stringify(computeAccountBalances_(ss)));
}
