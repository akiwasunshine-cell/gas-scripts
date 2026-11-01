// 売上データを月ごとに集計し、サマリーと棒グラフを出力するスクリプト

const SOURCE_SHEET_NAME = '売上データ';
const SUMMARY_SHEET_NAME = '月次サマリー';
const SUMMARY_HEADERS = ['月', '合計売上', '件数'];
const TRIGGER_FUNCTION_NAME = 'runDashboard';
const TRIGGER_HOUR = 9;

/**
 * メイン処理：売上データを集計し、月次サマリーとグラフを更新する。
 */
function runDashboard() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }
  const summarySheet = spreadsheet.getSheetByName(SUMMARY_SHEET_NAME)
    || spreadsheet.insertSheet(SUMMARY_SHEET_NAME);

  // 読み込みに失敗した場合に既存のサマリーを消さないよう、集計を終えてからクリアする
  const summaries = aggregateByMonth_(
    sourceSheet.getDataRange().getValues(),
    spreadsheet.getSpreadsheetTimeZone()
  );

  writeSummary_(summarySheet, summaries);
  drawChart_(summarySheet, summaries.length);
}

/**
 * 毎朝9時台にrunDashboardを実行するトリガーを登録する。
 * 実行のたびにトリガーが増えないよう、既存の登録を削除してから作成する。
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (trigger) {
      return trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME;
    })
    .forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone())
    .create();
}

/**
 * 売上データの全行を月ごとに集計し、古い月から順に返す。
 * 日付または金額が不正な行は集計せず、行番号をログに残す。
 */
function aggregateByMonth_(values, timeZone) {
  const totals = {};
  const skippedRows = [];

  values.forEach(function (row, index) {
    if (row.every(function (cell) { return cell === ''; })) {
      return;
    }
    const yearMonth = parseYearMonth_(row[0], timeZone);
    const amount = row[3];

    // 1行目の日付が読めない場合は見出し行とみなす
    if (!yearMonth && index === 0) {
      return;
    }
    if (!yearMonth || typeof amount !== 'number' || !isFinite(amount)) {
      skippedRows.push(index + 1);
      return;
    }

    const key = yearMonth.year * 100 + yearMonth.month;
    if (!totals[key]) {
      totals[key] = { year: yearMonth.year, month: yearMonth.month, total: 0, count: 0 };
    }
    totals[key].total += amount;
    totals[key].count += 1;
  });

  if (skippedRows.length > 0) {
    console.warn('集計から除外した行：' + skippedRows.join(', '));
  }

  return Object.keys(totals)
    .map(Number)
    .sort(function (a, b) { return a - b; })
    .map(function (key) { return totals[key]; });
}

/**
 * セルの値から年月を取り出す。日付型と「2026/01/05」形式の文字列に対応する。
 * 読み取れない場合はnullを返す。
 */
function parseYearMonth_(value, timeZone) {
  // 月の境界がずれないよう、日付型はスプレッドシートのタイムゾーンで解釈する
  const text = value instanceof Date && !isNaN(value.getTime())
    ? Utilities.formatDate(value, timeZone, 'yyyy/MM/dd')
    : String(value).trim();

  const matched = text.match(/^(\d{4})[\/-](\d{1,2})[\/-]\d{1,2}$/);
  if (!matched) {
    return null;
  }
  const month = Number(matched[2]);
  if (month < 1 || month > 12) {
    return null;
  }
  return { year: Number(matched[1]), month: month };
}

/**
 * 月次サマリーシートをクリアして集計結果を書き込む。
 */
function writeSummary_(sheet, summaries) {
  sheet.clear();

  const rows = summaries.map(function (summary) {
    return [summary.year + '年' + summary.month + '月', summary.total, summary.count];
  });
  const output = [SUMMARY_HEADERS].concat(rows);

  // 「2026年1月」が日付に自動変換されないよう、A列を書式なしテキストにしておく
  sheet.getRange(1, 1, output.length, 1).setNumberFormat('@');
  sheet.getRange(1, 1, output.length, SUMMARY_HEADERS.length).setValues(output);
  sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length).setFontWeight('bold');

  if (rows.length > 0) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0');
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('0');
  }
}

/**
 * 月次推移の棒グラフを作り直す。
 * シートをクリアしてもグラフは消えないため、重複しないよう先に削除する。
 */
function drawChart_(sheet, monthCount) {
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });

  if (monthCount === 0) {
    return;
  }

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, monthCount + 1, 2))
    .setNumHeaders(1)
    .setPosition(2, SUMMARY_HEADERS.length + 2, 0, 0)
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上' })
    .build();
  sheet.insertChart(chart);
}
