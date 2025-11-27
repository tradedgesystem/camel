/**
 * Google Backtesting Sheet Generator
 * MES-only (by default) with nicer UI:
 * - Custom menu
 * - Sidebar "builder" UI
 * - Styled Trades sheet (banding, colors, conditional formatting)
 * - Styled Summary sheet
 */

const DEFAULT_NUM_DAYS = 100;
const DEFAULT_TICKER = "MES";   // default ticker
const TRADE_RESULTS = [
  "🟩 Green (Win)",
  "🟥 Red (Loss)",
  "⬛ Breakeven",
  "🚫 No Trade",
  "❌ Did Not Trade Ticker Today"
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Backtesting Template')
    .addItem('Build / Reset (defaults)', 'buildTemplateWithDefaults')
    .addItem('Open Template Builder…', 'openTemplateBuilder')
    .addToUi();
}

/**
 * One-click version:
 * - Start date = Nov 24 of current year
 * - 100 trading days backward
 * - Ticker = MES
 */
function buildTemplateWithDefaults() {
  const ui = SpreadsheetApp.getUi();

  const button = ui.alert(
    'Rebuild template?',
    'This will DELETE existing "Trades" and "Summary" sheets and recreate them with default settings.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (button !== ui.Button.YES) return;

  const year = new Date().getFullYear();
  const startDate = new Date(year, 10, 24); // yyyy-11-24

  buildTemplate({
    startDate,
    numDays: DEFAULT_NUM_DAYS,
    ticker: DEFAULT_TICKER
  });

 * Generate structured day objects:
 * { startDate: Date, endDate: Date, label: "YYYY-MM-DD" }
 */
function generateTradingDays(startDate, daysToGenerate) {
  const days = [];
  let cursor = new Date(startDate);

  while (days.length < daysToGenerate) {
    const dow = cursor.getDay();
    if (dow !== 0 && dow !== 6) {
      days.push({
        startDate: new Date(cursor),
        endDate: new Date(cursor),
        label: Utilities.formatDate(
          cursor,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        )
      });
    }
    cursor.setDate(cursor.getDate() - 1);
  }

  return days;
}

/**
 * Trades sheet
 */
function createTradesSheet(tradingDays, ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName('Trades');
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet('Trades', 0);

  const headers = [
    'Date',
    'Ticker',
    'Time Taken',
    'Trade Result',
    'Points Risked',
    'Points Result',
    'Notes'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  const rows = [];
  tradingDays.forEach(day => {
    rows.push([day.startDate, ticker, '', '', '', '', '']);
  });

  const dataRange = sheet.getRange(2, 1, rows.length, headers.length);
  dataRange.setValues(rows);

  const lastRow = 1 + rows.length;

  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('hh:mm');

  sheet.autoResizeColumns(1, headers.length);

  sheet.getRange(1, 1, lastRow, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // Data validation
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TRADE_RESULTS, true)
    .setAllowInvalid(false)
    .build();

  const tradeResultRange = sheet.getRange(2, 4, rows.length, 1);
  tradeResultRange.setDataValidation(rule);

  // Conditional formatting
  let rules = sheet.getConditionalFormatRules();

  const addFormat = (match, bg, fg) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(match)
        .setBackground(bg)
        .setFontColor(fg)
        .setRanges([tradeResultRange])
        .build()
    );
  };

  addFormat("🟩 Green (Win)", '#d1fae5', '#065f46');
  addFormat("🟥 Red (Loss)", '#fee2e2', '#991b1b');
  addFormat("⬛ Breakeven", '#e5e7eb', '#111827');
  addFormat("🚫 No Trade", '#fef3c7', '#92400e');
  addFormat("❌ Did Not Trade Ticker Today", '#f3f4f6', '#4b5563');

  sheet.setConditionalFormatRules(rules);
}

/**
 * Summary sheet
 */
function createSummarySheet(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName('Summary');
  if (existing) ss.deleteSheet(existing);

  const summary = ss.insertSheet('Summary', 1);

  const metrics = [
    ['Metric', 'Value'],

    ['Total trading days',
      '=COUNTA(UNIQUE(FILTER(Trades!A2:A,Trades!A2:A<>"")))'],

    ['Number of no-trade days',
      '=LET(d,UNIQUE(FILTER(Trades!A2:A,Trades!A2:A<>"")),' +
      'SUM(MAP(d,LAMBDA(x,IF(SUMPRODUCT((Trades!A2:A=x)*' +
      '((Trades!D2:D="🟩 Green (Win)")+(Trades!D2:D="🟥 Red (Loss)")+(Trades!D2:D="⬛ Breakeven"))) = 0,1,0)))))'
    ],

    ['Total trades taken',
      '=COUNTIFS(Trades!D2:D,"🟩 Green (Win)")+' +
      'COUNTIFS(Trades!D2:D,"🟥 Red (Loss)")+' +
      'COUNTIFS(Trades!D2:D,"⬛ Breakeven")'
    ],

    ['Total points gained/lost',
      '=SUM(FILTER(Trades!F2:F,(Trades!D2:D="🟩 Green (Win)")+' +
      '(Trades!D2:D="🟥 Red (Loss)")+(Trades!D2:D="⬛ Breakeven")))'
    ],

    ['Average points risked',
      '=AVERAGE(FILTER(Trades!E2:E,Trades!E2:E<>"",' +
      '(Trades!D2:D="🟩 Green (Win)")+(Trades!D2:D="🟥 Red (Loss)")+' +
      '(Trades!D2:D="⬛ Breakeven")))'
    ],

    ['Win/Loss ratio (excl BE)',
      '=LET(w,COUNTIF(Trades!D2:D,"🟩 Green (Win)"),' +
      'l,COUNTIF(Trades!D2:D,"🟥 Red (Loss)"),IF(l=0,"N/A",w/l))'
    ],

    ['Average points per winning trade',
      '=AVERAGE(FILTER(Trades!F2:F,Trades!D2:D="🟩 Green (Win)"))'
    ],

    ['Average points per losing trade',
      '=AVERAGE(FILTER(Trades!F2:F,Trades!D2:D="🟥 Red (Loss)"))'
    ],

    ['Average Trade Time (hh:mm)',
      '=TEXT(AVERAGE(FILTER(Trades!C2:C,' +
      '(Trades!D2:D="🟩 Green (Win)")+(Trades!D2:D="🟥 Red (Loss)")+' +
      '(Trades!D2:D="⬛ Breakeven"))),"hh:mm")'
    ],

    ['Most traded ticker', ticker]
  ];

  summary.getRange(1, 1, metrics.length, 2).setValues(metrics);

  const header = summary.getRange('A1:B1');
  header
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff');

  summary.getRange(1, 1, metrics.length, 2)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  summary.autoResizeColumns(1, 2);

  summary.getRange('A18').setValue('Ticker Metrics').setFontWeight('bold');

  summary.getRange('A19:F19')
    .setValues([['Ticker','Wins','Losses','Total Trades','Win Rate','Total Points']])
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff');

  summary.getRange('A20').setValue(ticker);

  summary.getRange('B20').setFormula('=COUNTIFS(Trades!D:D,"🟩 Green (Win)")');
  summary.getRange('C20').setFormula('=COUNTIFS(Trades!D:D,"🟥 Red (Loss)")');
  summary.getRange('D20').setFormula(
    '=COUNTIFS(Trades!D:D,"🟩 Green (Win)")+' +
    'COUNTIFS(Trades!D:D,"🟥 Red (Loss)")+' +
    'COUNTIFS(Trades!D:D,"⬛ Breakeven")'
  );
  summary.getRange('E20').setFormula('=IF(D20=0,0,B20/D20)');
  summary.getRange('F20').setFormula(
    '=SUM(FILTER(Trades!F:F,(Trades!D:D="🟩 Green (Win)")+' +
    '(Trades!D:D="🟥 Red (Loss)")+(Trades!D:D="⬛ Breakeven")))'
  );

  summary.getRange('E20').setNumberFormat('0.0%');
  summary.getRange('B20:D20').setNumberFormat('0.00');
  summary.getRange('F20').setNumberFormat('0.00');

  summary.autoResizeColumns(1, 6);
}

/**
 * Autofill
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Trades') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (col !== 4 || row <= 1) return;

  const val = e.range.getValue();
  const nextRow = row + 1;
  if (nextRow > sheet.getLastRow()) return;

  const cell2 = sheet.getRange(nextRow, 4);

  if (val === "❌ Did Not Trade Ticker Today") {
    cell2.setValue("❌ Did Not Trade Ticker Today");
    return;
  }

  const realTrades = ["🟩 Green (Win)", "🟥 Red (Loss)", "⬛ Breakeven"];
  if (realTrades.includes(val)) {
    cell2.setValue("🚫 No Trade");
    return;
  }
}
