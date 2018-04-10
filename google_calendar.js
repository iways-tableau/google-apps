
/* 今月のカレンダーからイベントを取得する */
function getCalendar_test() {
  
  // フォルダ、ファイル関係の定義
  var targetFolderIds = ["スプレッドシートのフォルダIDを入力"];
  var targetFolder;
  var folderName;
  var objFiles;
  var objFile;
  var fileName;
  
  // スプレッドシート関係の定義
  var ss;
  var key;
  var sheets;
  var sheetId;
  
  
  for (var i = 0; i < targetFolderIds.length; i++) {
    // Idから対象フォルダの取得
    targetFolder = DriveApp.getFolderById(targetFolderIds[i]);
    folderName = targetFolder.getName();
    
    // 対象フォルダ以下のSpreadsheetを取得
    objFiles = targetFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    while (objFiles.hasNext()) {
      objFile = objFiles.next();
      fileName = objFile.getName();
      
      // Spreadsheetのオープン
      ss = SpreadsheetApp.openByUrl(objFile.getUrl());
      key = ss.getId();
      sheets = ss.getSheets();
    }
  }
  
  /*スプレッドシートをクリア*/
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getRange(1,1,400,5);
  range.clear();
  
  /*列名を入力*/
  var range = sheet.getRange("A1").setValue("No.");
  var range = sheet.getRange("B1").setValue("タイトル");
  var range = sheet.getRange("C1").setValue("開始時刻");
  var range = sheet.getRange("D1").setValue("終了時刻");
  var range = sheet.getRange("E1").setValue("所要時間");
  
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var no=1; //No
  
  var myCal=CalendarApp.getCalendarById('メールアドレスを入力'); //特定のIDのカレンダーを取得
  
  var date=new Date(); //対象月を指定
  var startDate=new Date(date); //取得開始日
  var endDate=new Date(date);
  endDate.setMonth(endDate.getMonth()+1);　//取得終了日
  
  var myEvents=myCal.getEvents(startDate,endDate); //カレンダーのイベントを取得
  
  /* イベントの数だけ繰り返してシートに記録 */
  for each(var evt in myEvents){
    mySheet.appendRow(
      [
        no, //No
        evt.getTitle(), //イベントタイトル
        evt.getStartTime(), //イベントの開始時刻
        evt.getEndTime(), //イベントの終了時刻
        "=INDIRECT(\"RC[-1]\",FALSE)-INDIRECT(\"RC[-2]\",FALSE)" //所要時間を計算
      ]
    );
    no++;
  }
}