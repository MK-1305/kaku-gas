// 送り先チェックシート①のデータをCSV①に転記する関数
function transferData1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("送り先チェックシート➀");
  const dest = ss.getSheetByName("CSV➀");

  // ソースシートの最終行を取得し、3行目からの行数を計算
  const lastRow = source.getLastRow();
  const numRows = lastRow - 3 + 1; // 3行目～最終行

  // D,E列（列番号4,5）の範囲を3行目から取得
  const deRange = source.getRange(3, 4, numRows, 2);
  const deValues = deRange.getValues();

  // G～AD列（列番号7～30、24列分）の範囲を3行目から取得
  const gadRange = source.getRange(3, 7, numRows, 24);
  const gadValues = gadRange.getValues();

  // 各行のデータを連結：D,E列の値 + 空セルを1つ + G～AD列の値
  const outputData = [];
  for (let i = 0; i < deValues.length; i++) {
    // 連結結果は [D,E] + [""] + [G～AD] → 合計2 + 1 + 24 = 27列
    outputData.push(deValues[i].concat([""]).concat(gadValues[i]));
  }

  // CSV①の2行目以降の全データを値のみクリア（書式は保持）
  dest
    .getRange(2, 1, dest.getMaxRows() - 1, dest.getMaxColumns())
    .clearContent();

  // 出力先シートCSV①に、A列2行目からoutputDataを書き込む
  dest
    .getRange(2, 1, outputData.length, outputData[0].length)
    .setValues(outputData);
}

// 送り先チェックシート➁のデータをCSV②に転記する関数
function transferData2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("送り先チェックシート➁");
  const dest = ss.getSheetByName("CSV➁");

  const lastRow = source.getLastRow();
  const numRows = lastRow - 3 + 1;

  const deRange = source.getRange(3, 4, numRows, 2);
  const deValues = deRange.getValues();

  const gadRange = source.getRange(3, 7, numRows, 24);
  const gadValues = gadRange.getValues();

  const outputData = [];
  for (let i = 0; i < deValues.length; i++) {
    outputData.push(deValues[i].concat([""]).concat(gadValues[i]));
  }

  // CSV②の2行目以降の全データを値のみクリア（書式は保持）
  dest
    .getRange(2, 1, dest.getMaxRows() - 1, dest.getMaxColumns())
    .clearContent();

  // 出力先シートCSV②に、A列2行目からoutputDataを書き込む
  dest
    .getRange(2, 1, outputData.length, outputData[0].length)
    .setValues(outputData);
}

// 両方の転記処理をまとめて実行する関数
function transferAll() {
  transferData1();
  transferData2();
}
