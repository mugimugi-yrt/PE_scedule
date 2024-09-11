// 作業フォルダのID指定 (利用表の置く場所指定のために使用)
var folderId = "folder ID is private.";

// 調整表シートの指定 (調整表シート = 回答収集中のスプレッドシート)
var coordinateSpreadsheetId = "sheet ID is private.";
var coordinateSpreadsheet = SpreadsheetApp.openById(coordinateSpreadsheetId);
var coordinateSheetName = "調整シート";
var coordinateSheet = coordinateSpreadsheet.getSheetByName(coordinateSheetName);  // セル情報はこれを使って取得

// フォーム受信時に実行される関数 (要フォーム受信トリガー)
function onFormSubmit(e) {
  var formResponses = e.values;

  // 提出団体リストの更新
  var num = updateChecklist(formResponses);
  if (num === -1) {
    Logger.log("error: 提出団体が工体連団体リストにありません。");
  }

  // 施設利用希望を調整表に記入
  var clubName = formResponses[2];
  var requestData = makeRequestData(formResponses);  // データ整形
  makeResearchSheet(clubName, requestData);

  // 決定版にプルダウンを埋め込む
  makePullDown();
}

// 提出リストを更新
function updateChecklist(data) {
  // 提出された団体の格納されている行を特定
  var clubName = data[2];
  var clubRow = clubnumCheck(clubName);
  if (clubRow === -1) { return -1; }  // 団体がリストに存在していない

  // チェックボックスを反転(trueにして提出していることを確認できるようにする)
  var checkBox = coordinateSheet.getRange(clubRow, 2);
  checkBox.setValue(true);

  // 利用の有無を利用状況に記入(利用する場合、背景色を黄色に変更)
  var usingRequest = data[3];
  var usingRequestCell = coordinateSheet.getRange(clubRow, 3);
  if (usingRequest === "利用する") {
    var clubNameCell = coordinateSheet.getRange(clubRow, 1);
    clubNameCell.setBackground("yellow");
  }
  usingRequestCell.setValue(usingRequest);

  // 備考を記入(利用が無い団体は理由を記入)
  var note = "";
  if      (usingRequest === "利用する") { note = note + data[9]; }
  else if (usingRequest === "利用しない") {
    note = note + data[10] + "。";
    note = note + data[11];
  }
  if (note != "") {
    var noteCell = coordinateSheet.getRange(clubRow, 4);
    noteCell.setValue(note);
  }

  // 提出団体数更新
  var submitNumCell = coordinateSheet.getRange("D21");
  var submitNum = parseInt(submitNumCell.getValue(), 10) + 1;
  submitNumCell.setValue(submitNum);
  submitNumCell.setHorizontalAlignment("left");
}

// 提出された団体の格納されている行を返す(col=1, row=23~23+団体数)
function clubnumCheck(name) {
  // 団体数チェック
  var clubNumCell = coordinateSheet.getRange("B21");
  var clubNum = parseInt(clubNumCell.getValue(), 10);

  // チェック記入
  for (var i = 23; i < 23 + clubNum; i++) {
    var namelistCell = coordinateSheet.getRange(i, 1);
    var namelist = namelistCell.getValue();
    if (name === namelist) { return i; }
  }
  return -1;  // ここにたどり着いている時点でExcelで問題あり
}

// フォーム回答から施設利用希望を表に合う形に整形
function makeRequestData(data) {
  // 返却リストを作成 (一番左が月曜日、一番右が日曜日)
  var returnData = [
    [false, false, false, false, false, false, false], // 体育館
    [false, false, false, false, false, false, false], // 小体育館
    [false, false, false, false, false, false, false], // 屋外コート
    [false, false, false, false, false, false, false], // 緑町グラウンド
    [false, false, false, false, false, false, false]  // 緑町テニスコート
  ];

  // フォーム回答から、施設利用希望データを抽出
  var formData = [
    data[4].split(','), // 体育館
    data[5].split(','), // 小体育館
    data[6].split(','), // 屋外コート
    data[7].split(','), // 緑町グラウンド
    data[8].split(',')  // 緑町テニスコート
  ];

  // 返却リストに利用希望の曜日をtrueに変更
  for (var i = 0; i < 5; i++) {
    if (formData[i].length !== 0) {
      for (var j = 0; j < formData[i].length; j++) {
        var request = formData[i][j].trim();
        if      (request === "月曜日") { returnData[i][0] = true; }
        else if (request === "火曜日") { returnData[i][1] = true; }
        else if (request === "水曜日") { returnData[i][2] = true; }
        else if (request === "木曜日") { returnData[i][3] = true; }
        else if (request === "金曜日") { returnData[i][4] = true; }
        else if (request === "土曜日") { returnData[i][5] = true; }
        else if (request === "日曜日") { returnData[i][6] = true; }
      }
    }
  }
  return returnData;
}

// 調整表の作成
function makeResearchSheet(name, data) {
  // 横向きに格納していく
  for (var i = 0; i < 5; i++) {
    for (var j = 0; j < 7; j++) {
      var request = data[i][j];
      if (request) {
        var usingCell = coordinateSheet.getRange(i + 3, j + 2);
        var using = usingCell.getValue()
        if      (using === "") { usingCell.setValue(name); }
        else if (using != "") {
          using = using + ", " + name;
          usingCell.setValue(using);
        }
      }
    }
  }
}

// 調整表から決定版にプルダウンを作成(row=3~7 to 12~16, col=2~8)
function makePullDown() {
  // 横向きで格納していく
  for (var i = 3; i <= 7; i++) {
    for (var j = 2; j <= 8; j++) {
      var researchCell = coordinateSheet.getRange(i, j);
      var research = researchCell.getValue();
      var menu = [];
      if      (research === "") { menu.push("使用禁止"); }
      else if (research != "")  {
        research = research + ", 使用禁止";
        menu = research.split(",").map(function(item) {return item.trim();});
      }
      var pulldownCell = coordinateSheet.getRange(i + 9, j);
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(menu).build();
      pulldownCell.setDataValidation(rule);      
    }
  }
}

// 利用表作成＆リセット (要変更時トリガー)
function onCheck(e) {
  var sheet = e.source.getActiveSheet(); // 編集されたシートを取得
  var range = e.source.getActiveRange();
  var col = range.getColumn();
  var row = range.getRow();

  // 編集されたシートが調整シートでなければ、実行終了(実行時間短縮)
  if (sheet.getName() !== coordinateSheetName) { return; }

  if (sheet.getName() === coordinateSheetName) {
    // 利用表作成(B17のチェックボックスがtrueなら)
    if      (row === 17 && col === 2) { makeUsingSheet(); }

    // コピー&リセット(B18のチェックボックスがtrueなら)
    else if (row === 18 && col === 2) { reset_copy(); }
  }
}

// 利用表の作成
function makeUsingSheet() {
  var checkBox = coordinateSheet.getRange("B17");
  var check = checkBox.getValue();

  if (check) {
    var year = parseInt(coordinateSheet.getRange("B10").getValue(), 10);
    var month = parseInt(coordinateSheet.getRange("D10").getValue(), 10);
    var message = String(year) + "年" + String(month) + "月の利用表を作成します。よろしいですか？(作成に約1分かかります)";
    var response = Browser.msgBox("完成！", message, Browser.Buttons.OK_CANCEL);
    if (response === "ok") {
      // 施設利用状況を日曜日から格納
      var usingData = getUsingData();

      // 利用表スプシの作成
      var num = makeReservingSpreadSheet(year, month, usingData);
      if (num === -1) { SpreadsheetApp.getUi().alert("日付データを取得できませんでした。半角入力などのチェックをしてください。"); }
    }
    if (response === "ok" && num === 0) { SpreadsheetApp.getUi().alert("施設利用表の作成が完了しました"); }
    checkBox.setValue(false);
  }
}

// (生データが多かったので分割) 決定版表から施設利用状況を抽出
function getUsingData() {
  // 体育館, 小体育館, 屋外コート, 緑町グラウンド, 緑町テニスコートの順番に格納
  var returnData = [[], [], [], [], []];
  var col = ["H", "B", "C", "D", "E", "F", "G"];
  var row = ["12", "13", "14", "15", "16"];
  for (var i = 0; i < row.length; i++) {
    for (var j = 0; j < col.length; j++) {
      var usingData = coordinateSheet.getRange(col[j] + row[i]).getValue();
      returnData[i].push(usingData);
    }
  }
  return returnData;
}

// (セル設定がかなりジカなので注意)利用表スプシの作成
function makeReservingSpreadSheet(y, m, data) {
  var whenData = String(y) + "年" + String(m) + "月";
  var whereData = ["体育館", "小体育館", "屋外コート", "緑町グラウンド", "緑町テニスコート"];
  var reservingSpreadSheet = SpreadsheetApp.create(whenData + "施設利用表");
  var reservingSpreadSheetId = reservingSpreadSheet.getId();
  DriveApp.getFileById(reservingSpreadSheetId).moveTo(DriveApp.getFolderById(folderId));

  // シート1を体育館シートとして作成
  var gymSheet = reservingSpreadSheet.getSheetByName('シート1');
  if (gymSheet) { gymSheet.setName(whereData[0]); }

  // カレンダーを作成(他シートはこれをコピーして使う)
  // 幅と高さ調整
  gymSheet.setRowHeight(1, 30);
  gymSheet.setRowHeight(2, 10);
  gymSheet.setRowHeights(3, 3, 22);
  for(var i = 1; i <= 6; i++) {
    for (var j = 1; j <= 5; j++) {
      var num = j % 5;
      if      (num === 1) { gymSheet.setRowHeight(5 * i + j, 22); }
      else if (num === 2 || num === 4) { gymSheet.setRowHeight(5 * i + j, 23); }
      else if (num === 3 || num === 0) { gymSheet.setRowHeight(5 * i + j, 25); }
    }
  }
  gymSheet.setRowHeight(36, 10);
  gymSheet.setRowHeight(37, 23);
  gymSheet.setColumnWidths(1, 7, 200);

  // カレンダーセル(枠線)の作成
  var singlestyle = SpreadsheetApp.BorderStyle.SOLID;  // 通常枠線設定
  var doublestyle = SpreadsheetApp.BorderStyle.DOUBLE; // 二重枠線設定
  var weekdayCell = ["A", "B", "C", "D", "E", "F", "G"];
  var weekday = ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"];
  for (var i = 0; i < 7; i++) {
    var weekLabelCell = gymSheet.getRange(weekdayCell[i] + "3");
    var weekCell = gymSheet.getRange(weekdayCell[i] + "3:" + weekdayCell[i] + "5");
    weekCell.setHorizontalAlignment("center");
    weekCell.setFontSize(12);
    weekCell.setBorder(true, true, true, true, false, false, "black", singlestyle);
    weekLabelCell.setFontWeight("bold");
    weekLabelCell.setValue(weekday[i]);
  }
  var dateData = getDateData(y, m);
  if (dateData.length === 0) { return -1; }
  for (var i = 0; i < dateData.length; i++) {
    for (var j = 0; j < dateData[i].length; j++) {
      var dayCellStartNum = 6 + (j * 5);
      var dayCellEndNum = dayCellStartNum + 4;
      var dayLabelCell = gymSheet.getRange(weekdayCell[i] + String(dayCellStartNum));
      var dayCell = gymSheet.getRange(weekdayCell[i] + String(dayCellStartNum) + ":" + weekdayCell[i] + String(dayCellEndNum))
      if (dateData[i][j] != 0) {
        dayLabelCell.setFontWeight("bold");
        dayLabelCell.setHorizontalAlignment("left");
        dayLabelCell.setFontSize(12);
        if      (i === 0) { dayLabelCell.setFontColor("red"); }
        else if (i === 6) { dayLabelCell.setFontColor("blue"); }
        dayLabelCell.setValue(dateData[i][j]);
      }
      dayCell.setBorder(true, true, true, true, false, false, "black", singlestyle);
    }
  }
  var separeteLineCell = gymSheet.getRange("A5:G5");
  separeteLineCell.setBorder(null, null, true, null, null, null, "black", doublestyle);

  // 備考記入欄作成
  var noteLineCell = gymSheet.getRange("A37:G40");
  var noteCell = gymSheet.getRange("A38:G40");
  noteCell.merge();
  noteCell.setVerticalAlignment("top");
  noteLineCell.setFontSize(12);
  noteLineCell.setBorder(true, true, true, true, false, false, "black", singlestyle);
  gymSheet.getRange("A37").setValue("備考");

  // 施設名(タイトル)記入 (他施設はコピー時に書き換え)
  gymSheet.getRange("C1:E1").merge();
  var titleCell = gymSheet.getRange("C1");
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setFontSize(17);
  titleCell.setValue(whenData + whereData[0] + "利用表")

  // コピーして施設別に作成
  var miniGymSheet = gymSheet.copyTo(reservingSpreadSheet);
  miniGymSheet.setName("小体育館");
  miniGymSheet.getRange("C1").setValue(whenData + whereData[1] + "利用表")
  var groundSheet = gymSheet.copyTo(reservingSpreadSheet);
  groundSheet.setName("屋外コート");
  groundSheet.getRange("C1").setValue(whenData + whereData[2] + "利用表")
  var midoriGroundSheet = gymSheet.copyTo(reservingSpreadSheet);
  midoriGroundSheet.setName("緑町グラウンド");
  midoriGroundSheet.getRange("C1").setValue(whenData + whereData[3] + "利用表")
  var midoriTennisSheet = gymSheet.copyTo(reservingSpreadSheet);
  midoriTennisSheet.setName("緑町テニスコート");
  midoriTennisSheet.getRange("C1").setValue(whenData + whereData[4] + "利用表")

  // 曜日データ格納
  Logger.log(data);
  var sheetList = [gymSheet, miniGymSheet, groundSheet, midoriGroundSheet, midoriTennisSheet];
  for (var i = 0; i < data.length; i++) {
    var targetSheet = sheetList[i];
    for (var j = 0; j < data[i].length; j++) {
      for (var k = 0; k < 6; k++) {
        var dayLabelCellNum = 6 + (k * 5);
        var dayLabelCell = targetSheet.getRange(weekdayCell[j] + String(dayLabelCellNum));
        // 数字部分に数字が記入されているところに利用予約を記入
        if (dayLabelCell.getValue() != "") {
          var usingClub = data[i][j];
          var usingCellNum = dayLabelCellNum + 4;
          var usingCell = targetSheet.getRange(weekdayCell[j] + String(usingCellNum));
          if      (data[i][j] === "使用禁止") {
            var coloringCell = targetSheet.getRange(weekdayCell[j] + String(dayLabelCellNum + 1) + ":" + weekdayCell[j] + String(usingCellNum));
            coloringCell.setBackground("#a3a3a3");
          }
          else if (usingClub === "") { usingClub = "(予約なし)"; }
          usingCell.setHorizontalAlignment("right");
          usingCell.setFontSize(14);
          usingCell.setValue(usingClub);
        }
      }
    }
  }

  // ここまで来たら正常に動作している
  return 0;
}

// 日付を曜日ごとのリストに割り振る
function getDateData(y, m) {
  var dayData = [[], [], [], [], [], [], []];  // とりあえずデータ格納用
  var endDay = 0;    // その月の最後の日の数を記録する用
  var returnData = [];  // 返却データ用(0=日曜日)
  if      (m != 2) {
    if      (m === 1 || m === 3 || m === 5 || m === 7 || m === 8 || m === 10 || m === 12) { endDay = 31; }
    else if (m === 4 || m === 6 || m === 9 || m === 11) { endDay = 30; }
  }
  else if (m === 2) {
    if      (y % 4 != 0) { endDay = 28; }
    else if (y % 4 === 0) {
      if      (y % 100 != 0 || y % 400 === 0) { endDay = 29; }
      else if (y % 100 === 0 || y % 400 != 0) { endDay = 28; }
    }
  }
  if (endDay === 0) { return dayData; }  // ここにハマったらエラー

  // とりあえず格納リストに格納
  for(var d = 1; d <= endDay; d++) { dayData[d % 7].push(d) }

  // y年m月1日の曜日を取得
  var first = new Date(y, m - 1, 1);
  var firstDayNum = first.getDay();

  // とりあえずリストを曜日順に格納
  // 計算してみると、((8-1日の曜日の数)%7)で日曜日がどこから始まるかわかる…っぽい
  var checkNum = 0;
  var inputNum = (8 - firstDayNum) % 7;
  for (var i = 0; i < 7; i++) {
    // 1日よりも前の日付の場合は、先頭に0を突っ込む
    if (inputNum === 1) { checkNum = 1; }
    if (checkNum === 0) { dayData[inputNum].unshift(0); }

    // カレンダー形みたいにするために、length 6以下のリストに0を突っ込む
    while (dayData[inputNum].length < 6) { dayData[inputNum].push(0); }
    returnData.push(dayData[inputNum]);
    inputNum = (inputNum + 1) % 7;
  }

  // 上の「0を突っ込む」処理がイマイチピンとこない場合は、下のdebug用logger使用
  // Logger.log(returnData);
  return returnData;
}

// リセットと作業シートをコピー
function reset_copy() {
  var checkBox = coordinateSheet.getRange("B18");
  var check = checkBox.getValue();

  if (check) {
    var response = Browser.msgBox("注意！", "シートをリセットします。よろしいですか？(コピーも同時に作成します)", Browser.Buttons.OK_CANCEL);
    if (response === "ok") {
      // コピーを作成(シート名は「年_月」)
      var year = coordinateSheet.getRange("B10").getValue();
      var month = coordinateSheet.getRange("D10").getValue();
      var copySheetName = String(year) + "_" + String(month);
      var copySheet = coordinateSheet.copyTo(coordinateSpreadsheet);
      copySheet.setName(copySheetName);
      copySheet.getRange("B18").setValue(false);

      // リセット
      var researchCell = coordinateSheet.getRange("B3:H7"); // 調整表
      researchCell.clearContent();
      var decideCell = coordinateSheet.getRange("B12:H16"); // 決定版
      decideCell.clearContent();
      decideCell.clearDataValidations();
      var clubNumCell = coordinateSheet.getRange("B21");
      var clubNum = parseInt(clubNumCell.getValue(), 10);
      var num = String(23 + clubNum - 1);
      var nameCell = coordinateSheet.getRange("A23:A" + num);
      nameCell.setBackground("white");
      var checkCell = coordinateSheet.getRange("B23:B" + num);
      checkCell.setValue(false);
      var usingCell = coordinateSheet.getRange("C23:C" + num);
      usingCell.clearContent();
      num = String(23 + clubNum);
      var noteCell = coordinateSheet.getRange("D23:D" + num);
      noteCell.clearContent();
      var submitNumCell = coordinateSheet.getRange("D21");
      submitNumCell.setValue(0);
      submitNumCell.setHorizontalAlignment("left");
    }
  }

  if (response === "ok") { SpreadsheetApp.getUi().alert("リセットが完了しました"); }
  checkBox.setValue(false);
}
