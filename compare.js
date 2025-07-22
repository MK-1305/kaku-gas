//20250414 setをpushに変更
function compareCsvSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const csv1Sheet = spreadsheet.getSheetByName("CSV➀");
  const csv2Sheet = spreadsheet.getSheetByName("CSV➁");
  const resultSheet = spreadsheet.getSheetByName("CSV結果");

  // CSV➀とCSV➁の全データを取得（Date型のフォーマットは formatDateValues 関数で変換）
  const csv1Data = formatDateValues(csv1Sheet.getDataRange().getValues());
  const csv2Data = formatDateValues(csv2Sheet.getDataRange().getValues());

  // CSV➀の1行目をヘッダーとして扱う（ヘッダーはそのまま出力する）
  const header = csv1Data.length > 0 ? csv1Data[0] : null;

  // CSV➀のヘッダー行の背景色を取得
  let headerBg = [];
  if (header) {
    headerBg = csv1Sheet.getRange(1, 1, 1, header.length).getBackgrounds();
  }

  // CSV➀とCSV➁のデータ（ヘッダー行を除く）を文字列化して配列に格納
  // 20250414-重複行を削除すると行分けて商品が同じ人がいると出力されないので、setではなくpushを使う
  const csv1Array = [];
  for (let i = 1; i < csv1Data.length; i++) {
    csv1Array.push(csv1Data[i].join(","));
  }

  const csv2Array = [];
  for (let i = 1; i < csv2Data.length; i++) {
    csv2Array.push(csv2Data[i].join(","));
  }

  // 結果用の配列 unmatchedData と行ごとの背景色 unmatchedColors を用意
  const unmatchedData = [];
  const unmatchedColors = [];

  // ヘッダー行はそのまま結果の先頭に追加
  if (header) {
    unmatchedData.push(header);
  }

  // CSV➀にあってCSV➁にない行 → 赤色（重複は削除せず、すべて保持）
  for (let rowString of csv1Array) {
    // csv2Array に同じ文字列が含まれているかチェック（includesを使用）
    if (!csv2Array.includes(rowString)) {
      const row = rowString.split(",");
      // ヘッダーの列数と合致しているかチェック
      if (row.length === header.length) {
        // すべてのセルが空の場合はスキップ
        if (row.every((cell) => cell.trim() === "")) {
          continue;
        }
        unmatchedData.push(row);
        unmatchedColors.push("#deb2af"); // 赤色系の指定色
      }
    }
  }

  // CSV➁にあってCSV➀にない行 → 青色
  for (let rowString of csv2Array) {
    if (!csv1Array.includes(rowString)) {
      const row = rowString.split(",");
      if (row.length === header.length) {
        if (row.every((cell) => cell.trim() === "")) {
          continue;
        }
        unmatchedData.push(row);
        unmatchedColors.push("#b0c4de"); // 青色系の指定色
      }
    }
  }

  // CSV結果シートの内容と書式を全てクリア（値と書式の両方をリセット）
  resultSheet.clearContents();
  resultSheet.clearFormats();

  // 結果データを書き出す（ヘッダーも含む）
  if (unmatchedData.length > 0) {
    resultSheet
      .getRange(1, 1, unmatchedData.length, header.length)
      .setValues(unmatchedData);
  }

  // ヘッダー行の背景色を反映（CSV➀のヘッダー背景色をそのまま）
  if (header) {
    resultSheet.getRange(1, 1, 1, header.length).setBackgrounds(headerBg);
  }

  // ヘッダー行は色付けしないので、unmatchedColorsは結果の2行目以降に対応する
  for (let i = 0; i < unmatchedColors.length; i++) {
    let color = unmatchedColors[i];
    resultSheet.getRange(i + 2, 1, 1, header.length).setBackground(color);
  }
}

/**
 * 2次元配列の中でDate型のデータを yyyy/MM/dd 形式に変換
 * @param {Array} data 2次元配列（スプレッドシートの getValues() の結果）
 * @returns {Array} フォーマット済みの2次元配列
 */
function formatDateValues(data) {
  return data.map((row) =>
    row.map((cell) => {
      if (cell instanceof Date) {
        return Utilities.formatDate(
          cell,
          Session.getScriptTimeZone(),
          "yyyy/MM/dd"
        );
      }
      return cell;
    })
  );
}
