/**
 * Google Apps Script to create a futures backtesting Google Sheets template.
 */
const NUM_DAYS = 100;
const TICKERS = ["MES", "M2K", "MYM", "MCL"];
const TRADE_RESULTS = ["🟩 Green (Win)", "🟥 Red (Loss)", "⬛ Breakeven", "🚫 No Trade"];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Backtesting Template')
    .addItem('Build/Reset Template', 'buildTemplate')
    .addToUi();
}

function buildTemplate() {
  const tradingDays = generateTradingDays(NUM_DAYS);
  createTradesSheet(tradingDays);
  createSummarySheet();
}

function resetTemplate() {
  buildTemplate();
}

function generateTradingDays(daysToGenerate) {
  const days = [];
  let cursor = new Date();
  while (days.length < daysToGenerate) {
    const day = new Date(cursor);
    const dayOfWeek = day.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      days.push(new Date(day.getFullYear(), day.getMonth(), day.getDate()));
    }
    cursor.setDate(cursor.getDate() - 1);
  }
  return days.reverse();
}

function createTradesSheet(tradingDays) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Trades';
  const existing = ss.getSheetByName(sheetName);
  if (existing) {
    ss.deleteSheet(existing);
  }
  const sheet = ss.insertSheet(sheetName, 0);
  const headers = ['Date', 'Ticker', 'Trade Result', 'Points Risked', 'Points Result', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  const rows = [];
  tradingDays.forEach((day) => {
    TICKERS.forEach((ticker) => {
      rows.push([day, ticker, '', '', '', '']);
    });
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.autoResizeColumns(1, headers.length);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TRADE_RESULTS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, rows.length, 1).setDataValidation(rule);

  sheet.setFrozenRows(1);
  return sheet;
}

function createSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Summary';
  const existing = ss.getSheetByName(sheetName);
  if (existing) {
    ss.deleteSheet(existing);
  }
  const summary = ss.insertSheet(sheetName, 1);

  const metrics = [
    ['Metric', 'Value'],
    ['Total trading days', '=COUNTA(UNIQUE(FILTER(Trades!A2:A,Trades!A2:A<>"")))'],
    ['Number of no-trade days', '=LET(d,UNIQUE(FILTER(Trades!A2:A,Trades!A2:A<>"")),SUM(MAP(d,LAMBDA(x,IF(SUMPRODUCT((Trades!A2:A=x)*(Trades!C2:C<>"🚫 No Trade"))=0,1,0)))))'],
    ['Total trades taken', '=COUNTIFS(Trades!C2:C,"<>🚫 No Trade",Trades!C2:C,"<>")'],
    ['Total points gained/lost', '=SUM(FILTER(Trades!E2:E,Trades!C2:C<>"🚫 No Trade"))'],
    ['Average points risked', '=AVERAGE(FILTER(Trades!D2:D,Trades!C2:C<>"🚫 No Trade",Trades!D2:D<>""))'],
    ['Win/Loss ratio (excl. BE)', '=LET(w,COUNTIF(Trades!C2:C,"🟩 Green (Win)"),l,COUNTIF(Trades!C2:C,"🟥 Red (Loss)"),IF(l=0,"N/A",w/l))'],
    ['Average points per winning trade', '=AVERAGE(FILTER(Trades!E2:E,Trades!C2:C="🟩 Green (Win)"))'],
    ['Average points per losing trade', '=AVERAGE(FILTER(Trades!E2:E,Trades!C2:C="🟥 Red (Loss)"))'],
    ['Most traded ticker', "=INDEX(SORT(QUERY(Trades!B2:C,'select B, count(C) where C <> \"🚫 No Trade\" group by B label count(C) \"\"',0),2,FALSE),1,1)"]
  ];
  summary.getRange(1, 1, metrics.length, 2).setValues(metrics);

  summary.getRange('A12').setValue('Ticker Metrics').setFontWeight('bold');
  summary.getRange('A13:F13').setValues([
    ['Ticker', 'Wins', 'Losses', 'Total Trades', 'Win Rate', 'Total Points']
  ]).setFontWeight('bold');

  TICKERS.forEach((ticker, index) => {
    const row = 14 + index;
    summary.getRange(row, 1).setValue(ticker);
    summary.getRange(row, 2).setFormula(`=COUNTIFS(Trades!B:B,"${ticker}",Trades!C:C,"🟩 Green (Win)")`);
    summary.getRange(row, 3).setFormula(`=COUNTIFS(Trades!B:B,"${ticker}",Trades!C:C,"🟥 Red (Loss)")`);
    summary.getRange(row, 4).setFormula(`=COUNTIFS(Trades!B:B,"${ticker}",Trades!C:C,"<>🚫 No Trade",Trades!C:C,"<>")`);
    summary.getRange(row, 5).setFormula(`=IF(D${row}=0,0,B${row}/D${row})`);
    summary.getRange(row, 6).setFormula(`=SUM(FILTER(Trades!E:E,Trades!B:B="${ticker}",Trades!C:C<>"🚫 No Trade"))`);
  });

  summary.autoResizeColumns(1, 6);
}
