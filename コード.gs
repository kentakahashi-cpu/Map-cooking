function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('お弁当配達ルートビルダー')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 💾 【新機能】画面から送られたJSONデータを解析し、顧客リストのスプシに自動保存する
function importJsonToMaster(jsonString) {
  if (!jsonString) return "エラー: JSONデータが空っぽです。";
  
  // 固定の顧客マスター用スプレッドシートの情報
  var MASTER_SHEET_ID = "1brsXsiqjNl1hO6Is1sp-eLXnVVNZg4kW0QqXSAN-6rA";
  var MASTER_SHEET_NAME = "顧客リスト";
  
  try {
    var data = JSON.parse(jsonString);
    var features = data.features || [];
    if (features.length === 0) return "エラー: 有効な地点データ（features）が見つかりませんでした。";
    
    var masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var masterSheet = masterSs.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "エラー: マスター側に「顧客リスト」という名前のシートが見つかりません。";
    
    // 現在登録されているマスターを読み込み（重複チェック用）
    var lastRow = masterSheet.getLastRow();
    var existingNames = [];
    if (lastRow >= 2) {
      var existingValues = masterSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existingValues.length; i++) {
        existingNames.push(existingValues[i][0].toString().trim());
      }
    }
    
    var newRows = [];
    var skipCount = 0;
    
    // JSONのデータを1件ずつ処理
    for (var j = 0; j < features.length; j++) {
      var feat = features[j];
      var props = feat.properties || {};
      var loc = props.location || {};
      var geom = feat.geometry || {};
      var coords = geom.coordinates || [0, 0];
      
      var name = loc.name ? loc.name.toString().trim() : "";
      var address = loc.address ? loc.address.toString().trim() : "";
      var lng = coords[0]; // JSONは[経度, 緯度]の順
      var lat = coords[1];
      
      if (!name) continue;
      
      // すでに同じ名前がスプシにある場合は上書きせずスキップ（重複防止）
      if (existingNames.indexOf(name) !== -1) {
        skipCount++;
        continue;
      }
      
      // [名前, 住所, 緯度, 経度] のセットを作成
      newRows.push([name, address, lat, lng]);
    }
    
    // 新しいデータがあればスプレッドシートの末尾に追記
    if (newRows.length > 0) {
      masterSheet.getRange(masterSheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
    }
    
    return "SUCCESS:" + newRows.length + "件の新しい顧客をスプシに保存しました！(重複スキップ: " + skipCount + "件)";
    
  } catch(e) {
    return "JSON解析エラー: " + e.message + "。データの形式を確認してください。";
  }
}

// 📅 指定された日付の注文データを、別スプシの「顧客リスト」と照合して同期する
function syncOrdersByDate(targetDateStr) {
  if (!targetDateStr) return "エラー: 日付が指定されていません。";
  
  var ORDERS_SHEET_ID = "12B6KRfnwwEHoagphWx5FG5lWMsnpD8o8aWlgsEI9qQE"; 
  var ORDERS_SHEET_NAME = "フォームの回答 1"; 
  var dateField = "お届け日";
  var nameField = "お名前";       
  var companyField = "会社名・所属";
  var timeField = "配達時間、配達場所の指定や、クーポンの使用、ご質問ご意見ご要望等ございましたら、お気軽にどうぞ";
  
  var MASTER_SHEET_ID = "1brsXsiqjNl1hO6Is1sp-eLXnVVNZg4kW0QqXSAN-6rA";
  var MASTER_SHEET_NAME = "顧客リスト";
  
  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var masterSheet = masterSs.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) return "エラー: マスター側に「顧客リスト」シートが見つかりません。";
    var masterLastRow = masterSheet.getLastRow();
    var masterRows = masterLastRow >= 2 ? masterSheet.getRange(2, 1, masterLastRow - 1, 4).getValues() : [];
    
    var orderSs = SpreadsheetApp.openById(ORDERS_SHEET_ID);
    var orderSheet = orderSs.getSheetByName(ORDERS_SHEET_NAME) || orderSs.getSheets()[0];
    var orderLastRow = orderSheet.getLastRow();
    
    if (orderLastRow < 2) return "注文データが空っぽです。";
    
    var orderHeaders = orderSheet.getRange(1, 1, 1, orderSheet.getLastColumn()).getValues()[0];
    var orderValues = orderSheet.getRange(2, 1, orderLastRow - 1, orderSheet.getLastColumn()).getValues();
    
    var dateIdx = orderHeaders.indexOf(dateField);
    var nameIdx = orderHeaders.indexOf(nameField);
    var compIdx = orderHeaders.indexOf(companyField);
    var timeIdx = orderHeaders.indexOf(timeField);
    
    if (dateIdx === -1 || nameIdx === -1) {
      return "エラー: 注文書に「お届け日」または「お名前」の列が見つかりません。";
    }
    
    var formattedTargetDate = targetDateStr.replace(/-/g, '/');
    var todaysOrders = [];
    
    for (var i = 0; i < orderValues.length; i++) {
      var row = orderValues[i];
      if (!row[dateIdx]) continue;
      
      var rawDate = row[dateIdx];
      var orderDateStr = "";
      if (rawDate instanceof Date) {
        orderDateStr = Utilities.formatDate(rawDate, "Asia/Tokyo", "yyyy/MM/dd");
      } else {
        orderDateStr = rawDate.toString().trim();
      }
      
      if (orderDateStr.indexOf(formattedTargetDate) !== -1 || formattedTargetDate.indexOf(orderDateStr) !== -1) {
        todaysOrders.push({
          name: row[nameIdx] ? row[nameIdx].toString().trim() : "",
          company: row[compIdx] ? row[compIdx].toString().trim() : "",
          time: (timeIdx !== -1 && row[timeIdx]) ? row[timeIdx].toString().trim() : ""
        });
      }
    }
    
    if (todaysOrders.length === 0) {
      return "指定された日付（" + formattedTargetDate + "）の注文データは見つかりませんでした。";
    }
    
    var outputValues = [];
    var matchCount = 0;
    
    for (var k = 0; k < todaysOrders.length; k++) {
      var order = todaysOrders[k];
      var matchedMaster = null;
      
      for (var m = 0; m < masterRows.length; m++) {
        var mName = masterRows[m][0].toString().trim();
        if (!mName) continue;
        
        if ((order.company && mName.indexOf(order.company) !== -1) || 
            (order.name && mName.indexOf(order.name) !== -1) ||
            (order.company && order.company.indexOf(mName) !== -1)) {
          matchedMaster = masterRows[m];
          break;
        }
      }
      
      if (matchedMaster) {
        var displayName = order.company ? order.company + " (" + order.name + ")" : order.name;
        outputValues.push([
          displayName,
          matchedMaster[1], 
          matchedMaster[2], 
          matchedMaster[3], 
          "all",            
          999,              
          order.time        
        ]);
        matchCount++;
      } else {
        if (!order.name && !order.company) continue;
        var displayName = order.company ? order.company + " (" + order.name + ") [要確認]" : order.name + " [要確認]";
        outputValues.push([displayName, "顧客リストに未登録です。住所を確認してください", 0, 0, "all", 999, order.time]);
      }
    }
    
    var currentSs = SpreadsheetApp.getActiveSpreadsheet();
    var destSheet = currentSs.getSheetByName("シート1") || currentSs.insertSheet("シート1");
    
    destSheet.clearContents();
    destSheet.getRange(1, 1, 1, 7).setValues([["名前", "住所", "緯度", "経度", "担当車両", "配達順", "配達時間"]]);
    
    if (outputValues.length > 0) {
      destSheet.getRange(2, 1, outputValues.length, 7).setValues(outputValues);
    }
    
    var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    PropertiesService.getScriptProperties().setProperty("LAST_IMPORTED_DATE", formattedTargetDate + " 分 (同期:" + now + ")");
    
    return "SUCCESS:" + todaysOrders.length + "件中、" + matchCount + "件を顧客リストと照合しました！";
    
  } catch(e) { return "エラー: " + e.message; }
}

function loadSavedData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
    var lastRow = sheet ? sheet.getLastRow() : 0;
    var scriptProperties = PropertiesService.getScriptProperties();
    var lastImportedDate = scriptProperties.getProperty("LAST_IMPORTED_DATE") || "";
    
    if (!sheet || lastRow < 2) return { places: [], lastDate: lastImportedDate };
    
    var range = sheet.getRange(2, 1, lastRow - 1, 7); 
    var values = range.getValues();
    var places = [];
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!row[0]) continue;
      places.push({
        id: 'place_' + i, title: row[0].toString(), address: row[1].toString(),
        lat: parseFloat(row[2]) || 0, lng: parseFloat(row[3]) || 0,
        car: row[4] ? row[4].toString() : 'all', order: row[5] ? parseInt(row[5]) : 999,
        time: row[6] ? row[6].toString() : ""
      });
    }
    return { places: places, lastDate: lastImportedDate };
  } catch(e) { return { places: [], lastDate: "エラー: " + e.message }; }
}

function saveDataToSheet(placesList) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
    if (!sheet) return "ERROR";
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 7).setValues([["名前", "住所", "緯度", "経度", "担当車両", "配達順", "配達時間"]]);
    if (placesList && placesList.length > 0) {
      var outputValues = [];
      for (var i = 0; i < placesList.length; i++) {
        var p = placesList[i];
        outputValues.push([p.title, p.address, p.lat, p.lng, p.car, p.order, p.time || ""]);
      }
      sheet.getRange(2, 1, outputValues.length, 7).setValues(outputValues);
    }
    return "SUCCESS";
  } catch(e) { return "エラー: " + e.message; }
}
