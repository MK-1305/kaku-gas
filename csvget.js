function generateShippingCheckSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName("明細一覧");
  const masterSheet = ss.getSheetByName("商品マスタ");
  const outputSheet = ss.getSheetByName("送り先チェックシート➁");

  // ★ 商品マスタの読み込み：各商品コードに対応するアルファベット（商品種別）をマップ作成
  const masterData = masterSheet.getDataRange().getValues();
  const productMap = {};
  for (let i = 0; i < masterData.length; i++) { // ヘッダー行は除く
    let code = masterData[i][0];         // A列：商品コード
    let alpha = masterData[i][2];          // C列：アルファベット
    if (code && alpha) {
      productMap[code] = alpha;
    }
  }

  // ★ 明細一覧シートの読み込み
  const detailData = detailSheet.getDataRange().getValues();
  // 使用する列（0始まりのインデックス）
  // C列：購入者名   → index 2
  // D列：受注番号   → index 3
  // G列：商品コード → index 6
  // N列：受注数     → index 13
  // P列：送り先名   → index 15

  // 受注番号ごとに集計するためのマップを作成
  // summaryMap: { 受注番号: { purchaser, recipient, items: { productType: 合計数量, ... } } }
  const summaryMap = {};
  for (let i = 1; i < detailData.length; i++) { // ヘッダーを除く
    const row = detailData[i];
    const orderNo = row[3];         // D列：受注番号
    const purchaser = row[2];       // C列：購入者名
    const recipient = row[15];      // P列：送り先名
    const productCode = row[6];     // G列：商品コード
    const quantity = Number(row[13]); // N列：受注数
    if (!orderNo || !productCode || !quantity) continue;
    // 商品マスタから、該当商品コードのアルファベット（商品種別）を取得
    const productType = productMap[productCode];
    if (!productType) continue; // 該当しなければスキップ
    if (!summaryMap[orderNo]) {
      summaryMap[orderNo] = {
        purchaser: purchaser,
        recipient: recipient,
        items: {}  // 各商品種別ごとの受注数合計を保持
      };
    }
    const items = summaryMap[orderNo].items;
    items[productType] = (items[productType] || 0) + quantity;
  }

  // ★ 出力する商品の種別は、固定のアルファベット（A～X：24文字）とする
  const productTypes = [];
  for (let i = 0; i < 24; i++) {
    productTypes.push(String.fromCharCode("A".charCodeAt(0) + i));
  }

  // ★ 出力データの作成（1行につき1注文）
  // 送り先チェックシートの更新対象は以下：
  // ・D列：購入者名、E列：送り先名（購入者と同じならE列は空）
  // ・G～AD列：各商品種別（A～X）の受注数（該当しなければ空）
  // ※ 出力は既存フォーマットを保持するため、値のみ更新する
  const output = [];
  const orders = Object.keys(summaryMap);
  for (let orderNo of orders) {
    const entry = summaryMap[orderNo];
    const purchaser = entry.purchaser;
    const recipient = entry.recipient;
    const row = [];
    // D列：購入者名
    row.push(purchaser);
    // E列：送り先名（購入者と同じなら空文字）
    row.push((recipient && recipient !== purchaser) ? recipient : "");
    // ※ 出力対象はD,E列とG～AD列の合計26列のうち、D,Eは先頭2列、G～ADは24列
    const prodData = [];
    for (let type of productTypes) {
      prodData.push(entry.items[type] || "");
    }
    row.push(...prodData);
    output.push(row);
  }

  // ★ 更新対象のセル
  // 既存のフォーマットを変更せず、値のみ更新するので clearContent() を使用
  // 更新する範囲：
  // - D,E列（列番号 4,5） : 2列分
  // - G～AD列（列番号 7～30） : 24列分
  // 更新開始行は3行目
  const startRow = 3;
  // 対象範囲の行数は、出力データの行数（注文数）とします
  const numOutputRows = output.length;

  // 対象範囲の値のみクリア（書式は保持）
  outputSheet.getRange(startRow, 4, outputSheet.getMaxRows() - startRow + 1, 2).clearContent();
  outputSheet.getRange(startRow, 7, outputSheet.getMaxRows() - startRow + 1, 24).clearContent();

  // ★ 書き込み
  // D,E列（列4,5）に、output の先頭2列を設定
  if (numOutputRows > 0) {
    const deData = output.map(row => row.slice(0, 2));
    outputSheet.getRange(startRow, 4, deData.length, 2).setValues(deData);
  }
  // G～AD列（列7～30）に、output の残りの列（24列分）を設定
  if (numOutputRows > 0) {
    const gadData = output.map(row => row.slice(2));
    Logger.log(JSON.stringify(gadData));
    outputSheet.getRange(startRow, 7, gadData.length, 24).setValues(gadData);
  }
}

function logProductMasterAlphabets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("商品マスタ");
  const masterData = masterSheet.getDataRange().getValues();
  const alphabets = [];
  // ヘッダー行を除いて、C列の値を収集
  for (let i = 0; i < masterData.length; i++) {
    const alpha = masterData[i][2]; // C列
    if (alpha) {
      alphabets.push(alpha);
    }
  }
  Logger.log("商品マスタに登録されているアルファベット： " + JSON.stringify(alphabets));
}

