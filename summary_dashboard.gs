/**
 * 売上データ集計ダッシュボード
 *
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートへ書き出す。
 * さらに月次推移を棒グラフで可視化する。
 * 毎朝9時の自動実行トリガーを setDailyTrigger() で登録できる。
 */

// シート名・列番号などの定数
var SOURCE_SHEET_NAME = '売上データ';   // 入力元シート
var SUMMARY_SHEET_NAME = '月次サマリー'; // 出力先シート

var COL_DATE = 0;   // A列：日付
var COL_PERSON = 1; // B列：担当者名（集計には未使用だが構造として保持）
var COL_PRODUCT = 2; // C列：商品名（同上）
var COL_AMOUNT = 3; // D列：金額

/**
 * メインの集計処理。
 * 売上データを月ごとに集計し、サマリーシートとグラフを更新する。
 */
function runDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  // 入力元シートが存在しない場合は処理を中断する
  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  // 月ごとの集計を実行する
  var summary = aggregateByMonth(sourceSheet);

  // サマリーシートを更新する
  var summarySheet = writeSummary(ss, summary);

  // 棒グラフを更新する
  updateChart(summarySheet, summary.length);
}

/**
 * 売上データシートを読み込み、月ごとに合計売上と件数を集計する。
 *
 * @param {Sheet} sourceSheet 売上データシート
 * @return {Array} [{label: '2026年1月', total: 金額, count: 件数}, ...] を月順に並べた配列
 */
function aggregateByMonth(sourceSheet) {
  // ヘッダー行（1行目）を除いたデータ範囲を取得する
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    return []; // データが無い場合は空配列を返す
  }

  var values = sourceSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // 月キー（例：2026-01）ごとに合計と件数を保持する
  var monthMap = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var dateValue = row[COL_DATE];
    var amount = row[COL_AMOUNT];

    // 日付が空、または日付として解釈できない行はスキップする
    var date = parseDate(dateValue);
    if (!date) {
      continue;
    }

    // 金額が数値でない行はスキップする
    if (typeof amount !== 'number' || isNaN(amount)) {
      continue;
    }

    // 月キーを「YYYY-MM」形式で作成する（並べ替え用）
    var year = date.getFullYear();
    var month = date.getMonth() + 1; // 0始まりのため +1
    var key = year + '-' + ('0' + month).slice(-2);

    // 初出の月はエントリを初期化する
    if (!monthMap[key]) {
      monthMap[key] = {
        label: year + '年' + month + '月', // 表示用ラベル
        total: 0,
        count: 0
      };
    }

    // 合計売上と件数を加算する
    monthMap[key].total += amount;
    monthMap[key].count += 1;
  }

  // 月キーを昇順に並べ替えて配列化する
  var keys = Object.keys(monthMap).sort();
  var result = [];
  for (var j = 0; j < keys.length; j++) {
    result.push(monthMap[keys[j]]);
  }

  return result;
}

/**
 * 様々な形式の日付入力を Date オブジェクトに変換する。
 *
 * @param {*} value セルの値（Date / 文字列 / その他）
 * @return {Date|null} 変換できた場合は Date、できなければ null
 */
function parseDate(value) {
  // すでに Date 型ならそのまま返す
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  // 文字列の場合は Date への変換を試みる
  if (typeof value === 'string' && value !== '') {
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  return null;
}

/**
 * 集計結果を「月次サマリー」シートへ書き込む。
 * シートが無ければ作成し、毎回クリアしてから書き直す。
 *
 * @param {Spreadsheet} ss 対象スプレッドシート
 * @param {Array} summary aggregateByMonth() の戻り値
 * @return {Sheet} 更新したサマリーシート
 */
function writeSummary(ss, summary) {
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  // サマリーシートが無ければ新規作成する
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }

  // 既存の内容を毎回クリアする
  summarySheet.clear();

  // ヘッダー行を書き込む（A列：月、B列：合計売上、C列：件数）
  var output = [['月', '合計売上', '件数']];

  // 集計結果を行データに変換する
  for (var i = 0; i < summary.length; i++) {
    output.push([summary[i].label, summary[i].total, summary[i].count]);
  }

  // まとめて書き込む
  summarySheet.getRange(1, 1, output.length, 3).setValues(output);

  // ヘッダー行を太字にして見やすくする
  summarySheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  return summarySheet;
}

/**
 * 月次推移の棒グラフを作成・更新する。
 * 既存のグラフは一度すべて削除してから作り直す。
 *
 * @param {Sheet} summarySheet サマリーシート
 * @param {number} dataRowCount 集計データの行数（ヘッダーを除く）
 */
function updateChart(summarySheet, dataRowCount) {
  // 既存のグラフをすべて削除する
  var charts = summarySheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    summarySheet.removeChart(charts[i]);
  }

  // データが無い場合はグラフを作成しない
  if (dataRowCount < 1) {
    return;
  }

  // 月（A列）と合計売上（B列）をグラフのデータ範囲とする（ヘッダー含む）
  var range = summarySheet.getRange(1, 1, dataRowCount + 1, 2);

  // 棒グラフを作成する
  var chart = summarySheet.newChart()
    .asColumnChart()
    .addRange(range)
    .setPosition(2, 5, 0, 0) // E列付近に配置する
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .build();

  summarySheet.insertChart(chart);
}

/**
 * 毎朝9時に runDashboard() を自動実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の同名トリガーを削除してから設定する。
 */
function setDailyTrigger() {
  // 既存の runDashboard 向けトリガーを削除する
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎日午前9時台に実行する時間主導型トリガーを作成する
  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}
