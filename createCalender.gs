// 入社する人のカレンダー登録
function createEvent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシートを取得
  const listSheet = ss.getSheetByName("カレンダー登録"); //シートを取得
  
  // カレンダー登録するための情報を連想配列に入れる
  const calendarinfo = getCdataInfo(listSheet);
  const count = Object.keys(calendarinfo).length
  
  // 1つずつカレンダー登録を行う
  const calendar = CalendarApp.getCalendarById("<<メールアドレス>>");
  for(i = 0 ; i < count ; i++ ){
    // 1人分のカレンダー情報を作成する
    const joindate = calendarinfo[i][0]
    const lname = calendarinfo[i][1]
    const fname = calendarinfo[i][2]
    
    const title = "入社イベント" + ":" + lname + fname;
    const date = joindate;
    const options = {
      description: "入社イベント" + ":" + lname + fname,
      guests: "<<メールアドレス>>", //実際の使用時はゲストアドレスを変更してください。
      sendInvites: true
    }
    // カレンダー登録実行
    const ret = calendar.createAllDayEvent(title, date, options);
    //エラーメッセージを取得し結果セルに記述する
    resultCalSpredSheet(i, listSheet, ret)
  }
}

// カレンダー登録するための情報を連想配列に入れる
function getCdataInfo(sheet) {  
  // 成功がついてない行番号をとる
  const lastARow = sheet.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  // 最後のカレンダー登録データ
  const lastBRow = sheet.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  // カラム数
  const last1Col = sheet.getRange("A1").getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  
  const cdatas = sheet.getRange(lastARow+1, 2, lastBRow-lastARow, last1Col-1).getValues();
  const cdata={};
  const count=cdatas.length;
  // 連想配列化
  for(var i=0;i<count;i++){
    cdata[i]=cdatas[i];
  }
  return cdata;
}

// 結果をスプレに入れる
function resultCalSpredSheet(i, sheet, ret){  
  const row = i + 3
  if (ret["error"]) {
    sheet.getRange(row,1).setValue(ret.error.message);
  } else {
    sheet.getRange(row,1).setValue("成功");
  }  
}