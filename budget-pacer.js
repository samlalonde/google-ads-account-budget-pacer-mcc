/******************************************************************************************
 * @name Budget Tracker | Monthly Spend  | MCC-Level 
 *
 * @overview
 * Creates pacing dashboards for accounts under a Google Ads Manager (MCC). For each client
 * account, the script generates a Google Sheet with:
 *   - Monthly budget pacing overview
 *   - Daily spend tables
 *   - Line graphs showing forecast vs. actual
 *   - Recommended new daily budgets to stay on track
 * Includes color-coded metrics (red/yellow/green) and a progress bar for quick triage. 
 * Ideal for small agencies or in-house teams managing multiple accounts.
 * Focus is only on Account-level budgets NOT for campaigns or adgroups.
 *
 * @instructions
 * 1) In your MCC: Tools & Settings → Bulk Actions → Scripts → New Script.
 * 2) Paste this file. Review the CONFIG section:
 *    - Set default time zone (default = GMT 0).
 *    - Configure per-account monthly budgets in the template sheet.
 *    - Ensure email + sheet permissions are enabled for your MCC.
 * 3) Authorize and Preview to verify logs and generated Sheets.
 * 4) Run the script.
 * 5) Go into the Google Sheet CONFIG and add the monthly budget to the Monthly Budget Column
 * 6) Re-run the script
 * 7) Schedule to run daily so pacing data stays fresh.
 *
 * @author Sam Lalonde
 * https://www.linkedin.com/in/samlalonde/ - sam@samlalonde.com
 *
 * @license
 * MIT — Free to use, modify, and distribute. See https://opensource.org/licenses/MIT
 *
 * @version
 * 1.0
 *
 * @changelog
 * - v1.0
 *   - Initial release with MCC-level budget pacing, daily spend tables, and forecast graphs.
 *   - Account-specific tabs include pacing % and recommended budget adjustments.
 *   - Supports automatic currency detection via getCurrencyCode().
 ******************************************************************************************/

// ===============================
// CONFIG
// ===============================

// Insert Google Sheet ID or URL below.
var USER_SHEET_URL  = 'https://docs.google.com/spreadsheets/d/INSERTSHEETID/'; 

// Use 'MCC' or 'AUTO' to use the account/MCC time zone.
var TIMEZONE = 'MCC';

// Lookback window for the weighted recent average (newest day highest weight)
var WMA_WINDOW_DAYS = 7;

// ===========================================================

var CONFIG_HEADERS = [
  'Account ID (digits only)',
  'Account Name',
  'Monthly Budget',
  'Include? (TRUE/FALSE)'
];

var OVERVIEW_HEADERS = [
  'Account Name','Account ID','Account Sheet',
  'Budget Cap','Spend to Date',
  'Progress (bar)',
  'Trend (vs Target)',
  'Pace Delta % (vs Target)',
  'Available Budget Remaining',
  'Days in Month','Days Elapsed',
  'Target Spend To Date',
  'Pace vs Target',
  'Percentage Budget Spent',
  'Projected EoM Spend',
  'Recommended Daily Spend to 100%',
  'Account Currency'
];

var PROP = { SPREADSHEET_ID: 'BUDGET_PACING_SPREADSHEET_ID' };

function main() {
  var tz = getTz_();
  bannerLog_('START RUN', {
    tz: tz,
    preview: AdWordsApp.getExecutionInfo().isPreview(),
    when: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss '(" + tz + ")'")
  });

  var ss = getUserSpreadsheet_();
  Logger.log('📄 Using Spreadsheet: ' + ss.getUrl());

  var shInstructions = getOrCreateSheet_(ss, 'Instructions');
  var shConfig       = getOrCreateSheet_(ss, 'Config');
  var shOverview     = getOrCreateSheet_(ss, 'Overview');

  ensureBaseTabOrder_(ss, shInstructions, shConfig, shOverview);

  writeInstructions_(shInstructions, tz);
  initializeConfigIfEmpty_(shConfig);
  applyConfigHeaderNotes_(shConfig);
  var seed = seedOrMergeConfigWithAccounts_(shConfig);
  Logger.log('🌱 Config seeded/merged — added: ' + seed.added + ', names updated: ' + seed.namesUpdated + ', total rows: ' + seed.totalRows);

  prepareOverview_(shOverview);

  var cfg = readConfig_(shConfig);
  Logger.log('⚙️  Config rows (valid & included): ' + cfg.rows.length);
  if (!cfg.rows.length) { Logger.log('ℹ️  Fill Config (budget + TRUE) and re-run.'); return; }

  var ids = cfg.rows.map(function(r){return r.accountId;});
  Logger.log('🔀 Processing ' + ids.length + ' account(s) in chunks of 50…');

  var overviewRows = [];
  var accountSheetNames = [];
  var totals = { processed:0, skipped:0, errors:0 };

  for (var i=0; i<ids.length; i+=50) {
    var it = MccApp.accounts().withIds(ids.slice(i, i+50)).get();
    while (it.hasNext()) {
      var acct = it.next();
      var acctId = acct.getCustomerId().replace(/-/g,'');
      var rowCfg = cfg.index[acctId];
      if (!rowCfg) { totals.skipped++; continue; }

      try {
        MccApp.select(acct);

        var currency  = AdsApp.currentAccount().getCurrencyCode();
        var acctName  = acct.getName();
        var budgetCap = rowCfg.monthlyBudget;
        var spendMtd  = AdsApp.currentAccount().getStatsFor('THIS_MONTH').getCost();
        var mCtx      = monthMeta_(new Date(), tz);
        var availRem  = Math.max(budgetCap - spendMtd, 0);

        var dailyActuals = getDailySpendThisMonth_();
        var perDayBase   = buildPerDayRows_(dailyActuals, budgetCap, mCtx);

        var wmaDaily = computeWmaDaily_(perDayBase, mCtx.daysElapsed, WMA_WINDOW_DAYS);
        var perDay   = applyWmaForecast_(perDayBase, mCtx, spendMtd, wmaDaily);

        var tabName = makeAccountTabName_(acctName, acctId);
        var sh = getOrCreateSheet_(ss, tabName);
        writeAccountSheet_(sh, {
          accountName: acctName,
          accountId: acctId,
          currency: currency,
          monthlyBudget: budgetCap,
          spendMtd: spendMtd,
          availableRemaining: availRem,
          daysInMonth: mCtx.daysInMonth,
          daysElapsed: mCtx.daysElapsed,
          perDay: perDay,
          tz: tz,
          updatedAt: new Date(),
          wmaDaily: wmaDaily
        });
        accountSheetNames.push(tabName);

        var targetToDate   = budgetCap * (mCtx.daysElapsed / mCtx.daysInMonth);
        var paceVsTarget   = spendMtd - targetToDate;
        var pctBudgetSpent = budgetCap > 0 ? (spendMtd / budgetCap) : 0;

        var remainingDays  = Math.max(mCtx.daysInMonth - mCtx.daysElapsed, 0);
        var projectedEom   = spendMtd + wmaDaily * remainingDays;

        var recDaily = remainingDays > 0 ? Math.max((budgetCap - spendMtd) / remainingDays, 0) : 0;

        var paceDeltaPct = targetToDate > 0 ? (spendMtd / targetToDate) - 1 : 0;
        var trendText    = trendLabel_(paceDeltaPct);
        var link = '=HYPERLINK("' + ss.getUrl() + '#gid=' + sh.getSheetId() + '","Open")';

        overviewRows.push([
          acctName, acctId, link,
          budgetCap, spendMtd,
          '', // SPARKLINE added after write
          trendText,
          paceDeltaPct,
          availRem,
          mCtx.daysInMonth, mCtx.daysElapsed,
          targetToDate,
          paceVsTarget,
          pctBudgetSpent,
          projectedEom,
          recDaily,
          currency
        ]);

        totals.processed++;
      } catch (e) {
        Logger.log('❌ Error processing ' + acctId + ': ' + e);
        totals.errors++;
      }
    }
  }

  writeOverview_(shOverview, overviewRows);
  orderClientSheetsByName_(ss, shOverview, accountSheetNames);

  bannerLog_('END RUN', totals);
}

/* ========================= Time Zone Helper ========================= */

function getTz_() {
  if (!TIMEZONE) return 'UTC';
  var t = ('' + TIMEZONE).trim();
  if (t.toUpperCase() === 'MCC' || t.toUpperCase() === 'AUTO') {
    return AdsApp.currentAccount().getTimeZone();
  }
  return t; // IANA/Apps Script tz like 'UTC', 'Europe/London'
}

/* ========================= Spreadsheet Helpers ========================= */

function getUserSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP.SPREADSHEET_ID);
  if (!id) {
    if (!USER_SHEET_URL || USER_SHEET_URL.indexOf('/d/') === -1) {
      throw new Error('Please paste a valid Google Sheet URL into USER_SHEET_URL at top of script.');
    }
    var m = USER_SHEET_URL.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!m) throw new Error('Could not parse spreadsheetId from URL.');
    id = m[1];
    props.setProperty(PROP.SPREADSHEET_ID, id);
  }
  return SpreadsheetApp.openById(id);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** Base tabs must be: Instructions, Config, Overview (left→right) */
function ensureBaseTabOrder_(ss, shInstructions, shConfig, shOverview) {
  ss.setActiveSheet(shOverview);     ss.moveActiveSheet(3);
  ss.setActiveSheet(shConfig);       ss.moveActiveSheet(2);
  ss.setActiveSheet(shInstructions); ss.moveActiveSheet(1);
  ss.setActiveSheet(shOverview);
}

function writeInstructions_(sheet, tz) {
  sheet.clear();
  var rows = [
    ['Budget Pacing — Instructions', ''],
    ['', ''],
    ['Step 1', 'Use the "Config" tab (between Instructions and Overview). Paste Account IDs, set budgets.'],
    ['Step 2', 'Include defaults to TRUE. Set to FALSE to exclude an account.'],
    ['Step 3', 'Budgets are numeric (no $/€). Currency is taken from each Ads account.'],
    ['Step 4', 'Re-run the script from the MCC. Overview and account tabs will refresh.'],
    ['', ''],
    ['Notes', 'Timezone used by this sheet: ' + tz + ' (change TIMEZONE at top). Window: THIS_MONTH.'],
    ['Forecasting', 'Projected values use a weighted recent average of the last ' + WMA_WINDOW_DAYS + ' days (newer days weighted higher).']
  ];
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.getRange(1,1).setFontWeight('bold').setFontSize(14);
  sheet.autoResizeColumns(1,2);
}

function initializeConfigIfEmpty_(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.clear();
    sheet.getRange(1,1,1,CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, CONFIG_HEADERS.length);
  }
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2,3,Math.max(1,sheet.getLastRow()-1),1).setNumberFormat('0.00');
  }
}

function applyConfigHeaderNotes_(sheet) {
  var notes = [
    'Digits only (no dashes). Example: 5529798336',
    'Optional label; auto-updated from Google Ads when names change.',
    'Numeric monthly budget in account currency. Example: 25000',
    'Defaults to TRUE. Set FALSE to exclude from Overview.'
  ];
  for (var c=0; c<CONFIG_HEADERS.length; c++) {
    sheet.getRange(1, c+1).setNote(notes[c]);
  }
}

/** Seed/merge Config. NEW rows default Include=TRUE. */
function seedOrMergeConfigWithAccounts_(sheet) {
  var last = sheet.getLastRow();
  var map = {};
  var namesUpdated = 0;
  if (last >= 2) {
    var data = sheet.getRange(2,1,last-1,CONFIG_HEADERS.length).getValues();
    for (var i=0;i<data.length;i++){
      var rowIndex = i+2;
      var id = (data[i][0]||'').toString().replace(/-/g,'').trim();
      if (!/^\d+$/.test(id)) continue;
      map[id] = { rowIndex: rowIndex, name: (data[i][1]||'').toString() };
    }
  }
  var toAppend = [];
  var it = MccApp.accounts().get();
  while (it.hasNext()) {
    var a = it.next();
    var id = a.getCustomerId().replace(/-/g,'');
    var name = a.getName();
    if (map[id]) {
      if (name && map[id].name !== name) {
        sheet.getRange(map[id].rowIndex, 2).setValue(name);
        namesUpdated++;
      }
    } else {
      toAppend.push([id, name, '', true]); // include TRUE by default
      map[id] = { rowIndex: null, name: name };
    }
  }
  if (toAppend.length) {
    var start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, toAppend.length, CONFIG_HEADERS.length).setValues(toAppend);
    sheet.getRange(start, 3, toAppend.length, 1).setNumberFormat('0.00');
  }
  sheet.autoResizeColumns(1, CONFIG_HEADERS.length);
  return { added: toAppend.length, namesUpdated: namesUpdated, totalRows: sheet.getLastRow() - 1 };
}

/* ========================= Overview ========================= */

function prepareOverview_(sheet) {
  sheet.clear();
  sheet.getRange(1,1,1,OVERVIEW_HEADERS.length).setValues([OVERVIEW_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.clearConditionalFormatRules();
}

function writeOverview_(sheet, rows) {
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2,1,rows.length,OVERVIEW_HEADERS.length).setValues(rows);

    // SPARKLINE progress bar in F (green bar over colored background per status)
    for (var i=0; i<rows.length; i++) {
      var r = i + 2;
      sheet.getRange(r, 6).setFormula(
        '=SPARKLINE(E' + r + ', {"charttype","bar"; "max", D' + r + '; "color1","#2ecc71"; "color2","#eaf3ec"})'
      );
    }

    // Formats
    [4,5,9,12,13,15,16].forEach(function(c){ sheet.getRange(2,c,rows.length,1).setNumberFormat('0.00'); });
    sheet.getRange(2,8,rows.length,1).setNumberFormat('0.00%');  // pace delta %
    sheet.getRange(2,14,rows.length,1).setNumberFormat('0.00%'); // % budget spent

    // Column widths
    var widths = [200,135,110,120,120,150,130,130,150,110,110,160,140,150,160,190,120];
    for (var c=1;c<=widths.length;c++) sheet.setColumnWidth(c, widths[c-1]);
  }

  // Conditional formatting
  var rules = [];

  // Pace vs Target (M) red/green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground('#FADBD8')
    .setRanges([sheet.getRange(2, 13, Math.max(rows.length,1), 1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground('#D5F5E3')
    .setRanges([sheet.getRange(2, 13, Math.max(rows.length,1), 1)]).build());

  // Trend (G) from H
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2<-0.05').setBackground('#FADBD8').setFontColor('#922B21')
    .setRanges([sheet.getRange(2,7,Math.max(rows.length,1),1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ABS($H2)<=0.05').setBackground('#E8F5E9').setFontColor('#1E8449')
    .setRanges([sheet.getRange(2,7,Math.max(rows.length,1),1)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2>0.05').setBackground('#FDEBD0').setFontColor('#AF601A')
    .setRanges([sheet.getRange(2,7,Math.max(rows.length,1),1)]).build());

  // NEW: Progress bar cell (F) background from pace delta H:
  // RED if |H| >= 10%, YELLOW if 5% < |H| < 10%, GREEN if |H| <= 5%
  var fRange = sheet.getRange(2,6,Math.max(rows.length,1),1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($H2<=-0.10,$H2>=0.10)')
    .setBackground('#FADBD8').setRanges([fRange]).build()); // red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR(AND($H2>-0.10,$H2<-0.05),AND($H2>0.05,$H2<0.10))')
    .setBackground('#FDEBD0').setRanges([fRange]).build()); // yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ABS($H2)<=0.05')
    .setBackground('#D5F5E3').setRanges([fRange]).build()); // green

  sheet.setConditionalFormatRules(rules);
}

/* ========================= Read Config ========================= */

function readConfig_(sheet) {
  var out = { rows: [], index: {} };
  var last = sheet.getLastRow();
  if (last < 2) return out;
  var data = sheet.getRange(2,1,last-1,CONFIG_HEADERS.length).getValues();
  var seen = {};
  for (var i=0;i<data.length;i++){
    var rawId = (data[i][0]||'').toString().trim();
    var acctId = rawId.replace(/-/g,'');
    var acctName = (data[i][1]||'').toString().trim();
    var monthlyBudget = Number(data[i][2]||0);

    var includeCell = data[i][3];
    // DEFAULT TRUE: only exclude when explicitly FALSE/"FALSE"
    var include = !(includeCell === false || includeCell === 'FALSE');

    if (!acctId || !/^\d+$/.test(acctId)) { if (rawId) Logger.log('⚠️  Bad CID: ' + rawId); continue; }
    if (seen[acctId]) { Logger.log('⚠️  Duplicate CID, keeping first: ' + acctId); continue; }
    if (!include || monthlyBudget <= 0) { continue; }

    seen[acctId]=true;
    var entry = { accountId: acctId, accountName: acctName, monthlyBudget: monthlyBudget };
    out.rows.push(entry); out.index[acctId]=entry;
  }
  return out;
}

/* ========================= Account Tabs & Forecast ========================= */

function makeAccountTabName_(name, acctId) {
  var clean = (name || '').replace(/[\[\]\:\?\*\/\\]/g,' ');
  if (clean.length > 80) clean = clean.substring(0,80).trim();
  return clean + ' - ' + acctId;
}

function writeAccountSheet_(sheet, ctx) {
  sheet.clear(); removeAllCharts_(sheet); sheet.clearConditionalFormatRules();

  var targetToDate = ctx.monthlyBudget * (ctx.daysElapsed / ctx.daysInMonth);
  var paceVsTarget = ctx.spendMtd - targetToDate;
  var pctBudgetSpent = ctx.monthlyBudget > 0 ? (ctx.spendMtd / ctx.monthlyBudget) : 0;

  var remainingDays = Math.max(ctx.daysInMonth - ctx.daysElapsed, 0);
  var projectedEom  = ctx.spendMtd + ctx.wmaDaily * remainingDays;
  var recDaily = (remainingDays > 0) ? Math.max((ctx.monthlyBudget - ctx.spendMtd) / remainingDays, 0) : 0;

  var updatedAtStr = Utilities.formatDate(ctx.updatedAt, ctx.tz, "yyyy-MM-dd HH:mm:ss '(" + ctx.tz + ")'");

  var kpis = [
    ['Last Updated', updatedAtStr],
    ['Account Name', ctx.accountName],
    ['Account ID', ctx.accountId],
    ['Account Currency', ctx.currency],
    ['Budget Cap', ctx.monthlyBudget],
    ['Spend to Date', ctx.spendMtd],
    ['Available Budget Remaining', ctx.availableRemaining],
    ['Days in Month', ctx.daysInMonth],
    ['Days Elapsed', ctx.daysElapsed],
    ['Target Spend To Date', targetToDate],
    ['Pace vs Target', paceVsTarget],
    ['Percentage Budget Spent', pctBudgetSpent],
    ['Projected EoM Spend', projectedEom],
    ['Recommended Daily Spend to 100%', recDaily],
    ['Recent Daily Avg (last ' + WMA_WINDOW_DAYS + ' d)', ctx.wmaDaily]
  ];
  sheet.getRange(1,1,kpis.length,2).setValues(kpis);
  if (kpis.length >= 2) sheet.getRange(2,2,kpis.length-1,1).setNumberFormat('0.00');
  sheet.getRange(12,2).setNumberFormat('0.00%');

  // Traffic-light conditional formatting for KPI cells
  var rules = [];
  var rTarget = sheet.getRange(10,2,1,1);
  var rPace   = sheet.getRange(11,2,1,1);
  var rPct    = sheet.getRange(12,2,1,1);
  var rProj   = sheet.getRange(13,2,1,1);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B$10>0, ABS($B$11)/$B$10<=0.05)').setBackground('#D5F5E3').setRanges([rTarget]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B$10>0, ABS($B$11)/$B$10>0.05, ABS($B$11)/$B$10<=0.10)').setBackground('#FDEBD0').setRanges([rTarget]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($B$10=0, ABS($B$11)/$B$10>0.10)').setBackground('#FADBD8').setRanges([rTarget]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B$11>0.05*$B$10').setBackground('#D5F5E3').setRanges([rPace]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ABS($B$11)<=0.05*$B$10').setBackground('#FDEBD0').setRanges([rPace]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B$11<-0.05*$B$10').setBackground('#FADBD8').setRanges([rPace]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B$12>=0.95,$B$12<=1.05)').setBackground('#D5F5E3').setRanges([rPct]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR(AND($B$12>=0.90,$B$12<0.95),AND($B$12>1.05,$B$12<=1.10))').setBackground('#FDEBD0').setRanges([rPct]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($B$12<0.90,$B$12>1.10)').setBackground('#FADBD8').setRanges([rPct]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B$5>0, ABS($B$13-$B$5)/$B$5<=0.05)').setBackground('#D5F5E3').setRanges([rProj]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B$5>0, ABS($B$13-$B$5)/$B$5>0.05, ABS($B$13-$B$5)/$B$5<=0.10)').setBackground('#FDEBD0').setRanges([rProj]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($B$5=0, ABS($B$13-$B$5)/$B$5>0.10)').setBackground('#FADBD8').setRanges([rProj]).build());

  sheet.setConditionalFormatRules(rules);

  // Chart anchored at E1
  var CHART_ANCHOR_ROW = 1, CHART_ANCHOR_COL = 5, CHART_HEIGHT_ROWS = 18;
  var chartBottomRow = CHART_ANCHOR_ROW + CHART_HEIGHT_ROWS;

  var startRow = Math.max(kpis.length + 3, chartBottomRow + 2);
  var headers = [
    'Date','Cost (Day)','Cumulative Spend','Target Daily Spend',
    'Cumulative Forecast','Daily Gap (vs Target Daily)','Cumulative Gap (vs Target)',
    'Running Pace %','Projected EoM Spend','Recommended Daily Budget'
  ];
  sheet.getRange(startRow,1,1,headers.length).setValues([headers]).setFontWeight('bold');

  if (ctx.perDay.length) {
    var vals = ctx.perDay.map(function(r){
      return [
        r.date, r.cost, r.cumSpend, r.targetDaily,
        r.cumForecastWma, r.gap, r.cumGap,
        r.runningPacePct, r.projectedEomWmaAtDay, r.recDaily
      ];
    });
    sheet.getRange(startRow+1,1,vals.length,headers.length).setValues(vals);

    sheet.getRange(startRow+1,1,vals.length,1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(startRow+1,2,vals.length,6).setNumberFormat('0.00');
    sheet.getRange(startRow+1,9,vals.length,2).setNumberFormat('0.00');
    sheet.getRange(startRow+1,8,vals.length,1).setNumberFormat('0.00%');

    var widths = [110,110,140,150,170,160,170,130,180,180];
    for (var c=1;c<=widths.length;c++) sheet.setColumnWidth(c, widths[c-1]);
    sheet.setFrozenRows(0);

    var ticks = buildWeeklyTicksFromPerDay_(ctx.perDay);

    var rowsWithHeader = vals.length + 1;
    var chart = sheet.newChart()
      .asLineChart()
      .addRange(sheet.getRange(startRow, 1, rowsWithHeader, 1)) // Date
      .addRange(sheet.getRange(startRow, 3, rowsWithHeader, 1)) // Cumulative Spend
      .addRange(sheet.getRange(startRow, 5, rowsWithHeader, 1)) // Cumulative Forecast
      .setPosition(CHART_ANCHOR_ROW, CHART_ANCHOR_COL, 0, 0)    // E1
      .setOption('title','Pacing — Spend vs Forecast (This Month)')
      .setOption('legend',{ position:'right' })
      .setOption('useFirstColumnAsDomain', true)
      .setOption('useFirstRowAsHeaders', true)
      .setOption('series', {
        0: { labelInLegend: 'Cumulative Spend' },
        1: { labelInLegend: 'Cumulative Forecast' }
      })
      .setOption('width', 720)
      .setOption('height', 300)
      .setOption('hAxis', { format: 'MMM d', ticks: ticks })
      .build();
    sheet.insertChart(chart);
  }
}

/* ========================= Dates & Math ========================= */

function monthMeta_(today, tz) {
  var y  = Number(Utilities.formatDate(today, tz, 'yyyy'));
  var m0 = Number(Utilities.formatDate(today, tz, 'M')) - 1;
  var first = new Date(y, m0, 1);
  var next  = new Date(y, m0 + 1, 1);
  var end   = new Date(next.getTime() - 24*60*60*1000);
  var daysInMonth = end.getDate();
  var daysElapsed = Math.min(Number(Utilities.formatDate(today, tz, 'd')), daysInMonth);
  return { tz: tz, y: y, m0: m0, first: first, daysInMonth: daysInMonth, daysElapsed: daysElapsed };
}

function getDailySpendThisMonth_() {
  var awql = 'SELECT Date, Cost FROM ACCOUNT_PERFORMANCE_REPORT DURING THIS_MONTH';
  var rows = [];
  var report = AdsApp.report(awql);
  var it = report.rows();
  while (it.hasNext()) {
    var r = it.next();
    var d = r['Date'].split('-');
    rows.push({ date: new Date(+d[0], +d[1]-1, +d[2]), cost: parseFloat(r['Cost']) || 0 });
  }
  rows.sort(function(a,b){ return a.date - b.date; });
  return rows;
}

/** Build base per-day series across entire month (actuals + zeros for future) */
function buildPerDayRows_(daily, monthlyBudget, mCtx) {
  var byDay = {};
  for (var i=0;i<daily.length;i++) {
    var day = daily[i].date.getDate();
    byDay[day] = (byDay[day] || 0) + (daily[i].cost || 0);
  }
  var targetDaily = mCtx.daysInMonth > 0 ? (monthlyBudget / mCtx.daysInMonth) : 0;
  var perDay = [], cumSpend = 0;

  for (var d=1; d<=mCtx.daysInMonth; d++) {
    var cost = byDay[d] || 0;
    cumSpend += cost;

    var cumTarget = targetDaily * d;
    var gap       = cost - targetDaily;
    var cumGap    = cumSpend - cumTarget;
    var runningPacePct = monthlyBudget > 0 ? (cumSpend / monthlyBudget) : 0;

    perDay.push({
      date: new Date(mCtx.y, mCtx.m0, d),
      cost: cost,
      cumSpend: cumSpend,
      targetDaily: targetDaily,
      cumTarget: cumTarget,
      gap: gap,
      cumGap: cumGap,
      runningPacePct: runningPacePct,
      recDaily: (mCtx.daysInMonth - d > 0)
        ? Math.max((monthlyBudget - cumSpend) / (mCtx.daysInMonth - d), 0)
        : 0
    });
  }
  return perDay;
}

/** Weighted recent average of last N ACTUAL days; newest has highest weight. */
function computeWmaDaily_(perDay, daysElapsed, window) {
  var n = Math.min(window, Math.max(daysElapsed, 0));
  if (n <= 0) return 0;
  var sumW = 0, sumWX = 0;
  for (var i=0; i<n; i++) {
    var idx = daysElapsed - i; // 1-based day
    if (idx <= 0) break;
    var cost = perDay[idx - 1].cost || 0;
    var w = n - i; // weights 1..n (newest highest)
    sumW += w; sumWX += w * cost;
  }
  return sumW ? (sumWX / sumW) : 0;
}

/** Extend forecast cumulatively for charting + per-row projected EoM. */
function applyWmaForecast_(perDay, mCtx, spendMtd, wmaDaily) {
  for (var d=1; d<=perDay.length; d++) {
    perDay[d-1].cumForecastWma = (d <= mCtx.daysElapsed)
      ? perDay[d-1].cumSpend
      : spendMtd + wmaDaily * (d - mCtx.daysElapsed);
    perDay[d-1].projectedEomWmaAtDay =
      perDay[d-1].cumSpend + wmaDaily * (mCtx.daysInMonth - d);
  }
  return perDay;
}

/* ========================= Sheet Ordering ========================= */

function orderClientSheetsByName_(ss, overviewSheet, names) {
  if (!names || !names.length) return;
  var unique = {}; names.forEach(function(n){ unique[n]=true; });
  var clientSheets = ss.getSheets().filter(function(sh){ return unique[sh.getName()]; });
  if (!clientSheets.length) return;

  clientSheets.sort(function(a,b){
    var A=a.getName().toLowerCase(), B=b.getName().toLowerCase();
    return A<B?-1:A>B?1:0;
  });

  var overviewPos = getSheetPosition_(ss, overviewSheet); // 1-based
  for (var k=0; k<clientSheets.length; k++) {
    var sh = clientSheets[k];
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(overviewPos + 1 + k);
  }
  ss.setActiveSheet(overviewSheet);
}

function getSheetPosition_(ss, sheet) {
  var arr = ss.getSheets();
  for (var i=0;i<arr.length;i++) if (arr[i].getSheetId() === sheet.getSheetId()) return i+1;
  return 1;
}

/* ========================= Utilities ========================= */

function trendLabel_(paceDeltaPct) {
  var abs = Math.abs(paceDeltaPct), pct = Math.round(abs * 100);
  if (abs <= 0.05) return 'On Target';
  return (paceDeltaPct < 0) ? ('Under ' + pct + '%') : ('Over ' + pct + '%');
}

function buildWeeklyTicksFromPerDay_(perDay) {
  if (!perDay || !perDay.length) return [];
  var y = perDay[0].date.getFullYear(), m0 = perDay[0].date.getMonth();
  var lastDay = perDay[perDay.length - 1].date.getDate();
  var ticks = [];
  for (var d=1; d<=lastDay; d+=7) ticks.push(new Date(y, m0, d));
  if (ticks[ticks.length - 1].getDate() !== lastDay) ticks.push(new Date(y, m0, lastDay));
  return ticks;
}

function removeAllCharts_(sheet) {
  var charts = sheet.getCharts();
  for (var i=0;i<charts.length;i++) sheet.removeChart(charts[i]);
}

function bannerLog_(title, dataObj) {
  var border = Array(70).join('=');
  Logger.log(border);
  Logger.log('[' + title + ']');
  if (dataObj) for (var k in dataObj) if (dataObj.hasOwnProperty(k)) Logger.log(' - ' + k + ': ' + dataObj[k]);
  Logger.log(border);
}
