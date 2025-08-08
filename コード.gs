/**
 * このスプレッドシートを操作の対象として設定
 * もし実際のシート名が違う場合は、"シート1"の部分を修正してください
 */
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");

/**
 * Webページ(index.html)を表示するためのメイン関数
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

/**
 * 在庫データをすべて取得する関数
 */
function getInventory() {
  if (sheet.getLastRow() < 2) {
    return [];
  }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const inventoryList = data.map(row => {
    const item = {};
    headers.forEach((header, index) => {
      // 日付の場合は、読みやすい形式に変換
      if (header === 'last_updated' && row[index]) {
        item[header] = new Date(row[index]).toLocaleString();
      } else {
        item[header] = row[index];
      }
    });
    return item;
  });
  return inventoryList;
}

/**
 * 新しい商品を追加する関数
 */
function addItem(item) {
  const timestamp = new Date();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // IDを計算
  const ids = sheet.getRange("A2:A").getValues().flat().filter(id => typeof id === 'number');
  const maxId = ids.length > 0 ? Math.max(...ids) : 0;
  const newId = maxId + 1;
  
  // ヘッダーの順番に合わせて行データを作成
  const newRow = headers.map(header => {
    switch(header) {
      case 'id': return newId;
      case 'product_name': return item.product_name || '';
      case 'quantity': return item.quantity || 0;
      case 'last_updated': return timestamp;
      case 'category': return item.category || '';
      case 'notes': return item.notes || '';
      default: return '';
    }
  });

  sheet.appendRow(newRow);
  return `「${item.product_name}」を追加しました。`;
}

/**
 * 在庫情報を更新する関数
 */
function updateItem(item) {
  const timestamp = new Date();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  // findIndexは0から始まるので、実際の行番号は+2する(ヘッダー分+1、0-index分+1)
  const rowIndex = data.findIndex(row => row[0] == item.id);
  
  if (rowIndex !== -1) {
    const rowNum = rowIndex + 1;
    // 各列を更新
    sheet.getRange(rowNum, headers.indexOf('quantity') + 1).setValue(item.quantity);
    sheet.getRange(rowNum, headers.indexOf('category') + 1).setValue(item.category);
    sheet.getRange(rowNum, headers.indexOf('notes') + 1).setValue(item.notes);
    sheet.getRange(rowNum, headers.indexOf('last_updated') + 1).setValue(timestamp);
    return `ID: ${item.id} の商品を更新しました。`;
  }
  return "エラー: 更新対象の商品が見つかりません。";
}

/**
 * 在庫を削除する関数
 */
function deleteItem(id) {
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] == id);
  
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
    return `ID: ${id} の商品を削除しました。`;
  }
  return "エラー: 対象の商品が見つかりません。";
}
