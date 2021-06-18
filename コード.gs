/**
 * 全MLとユーザ一覧を取得
 */
function getGroupAddressAndUsers()
{
  // アクティブなシートを設定
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0]
  var pageToken, prev_pageToken;
  //スクリプト情報を取得
  var Properties = PropertiesService.getScriptProperties();
  //現在のページトークンを取得。無かったら初期化
  var pageToken = Properties.getProperty("task_page");
  if(pageToken == 'undefined') {
    pageToken = null;
  }

  var rows = [];  
  var starttime = new Date(); //時刻格納用の変数
  //タイトルを設定
  sheet.getRange(1, 1, 1, 6).setValues([["メールアドレス", "説明", "名前", "メンバー数", "メンバーのアドレス", "メンバーのロール"]]);

  do {
    //現在時刻と開始時刻の差分を取得
    var nowtime = new Date();
    var nowdiff = parseInt((nowtime.getTime()-starttime.getTime())/60000);

    // ”Admin Directory API”からグループ一覧を取得
    all_groups_page = AdminDirectory.Groups.list({
      domain: 'uluru.jp', //使用しているドメイン
      maxResults: 20,      //取得する件数
      pageToken: pageToken //トークン
    });
    var groups = all_groups_page.groups;

    //取得するグループが無い場合は終了
    if (!groups) {
      return;
    }

    // グループに所属するメンバーの取得し、配列に保存
    groups.forEach(function(group, i) {
      var members = AdminDirectory.Members.list(group.email).members;
      rows = addMemberRow(rows, group, members);
    })

    //現在時刻と開始時刻を比較して、差分が4分を超えていたら処理中断
    //（5分を超える処理はタイムアウトエラーになるため）
    if(nowdiff >= 4){

      // 配列の内容をシートへ書き込み
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
      rows = [];

      //次の処理のトリガーを設置、プロパティに現在のトークンを登録しプログラムを停止
      Properties.setProperty("task_page",pageToken);
      setTrigger();

      return;
    }

    //新しいトークンを取得
    prev_pageToken = pageToken;
    pageToken = all_groups_page.nextPageToken;            
  } while (pageToken);  

  //トークンに変化がない場合、完了
  if(prev_pageToken == pageToken) {    
    //プロジェクトトリガーを全削除
    deleteTrigger();
    //トリガー用変数の初期化
    Properties.setProperty("task_page",'');

    // シートに書き込み
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);

  //5分を超えずに終わった場合
  } else {
    // シートに書き込み
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    rows = [];
  }
  return;
}

//メンバーの配列データを返す
function addMemberRow(rows, group, members) {
  // グループに所属するメンバーの取得

  var nowtime = new Date();//////////////////////

  //メンバーが居なかった場合
  if(members == undefined || members == 0){
    var cols = [];
    cols.push(group.email);
    cols.push(group.description);
    cols.push(group.name);
    cols.push(group.directMembersCount);
    cols.push("-----");
    cols.push("-----");

    rows.push(cols);
  //メンバーが居た場合
  } else {
    for (var j = 0; j < members.length; j++){
      var cols = [];      
      cols.push(group.email);
      cols.push(group.description);
      cols.push(group.name);
      cols.push(group.directMembersCount);
      cols.push(members[j].email);
      cols.push(members[j].role);
      rows.push(cols);

    }
  }
  return rows;
}

//トリガーを全削除する関数
function deleteTrigger() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < allTriggers.length; i++) {
      ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

//トリガーを設置する関数（getGroupAddressAndUsers 関数を実行予定として登録）
function setTrigger(){
   deleteTrigger();
   var triggerman = ScriptApp.newTrigger("getGroupAddressAndUsers")
                   .timeBased()
                   .everyMinutes(5)
                   .create();
}

//トリガーをリセット
function resetTrigger() {
  var Properties = PropertiesService.getScriptProperties();

  //プロジェクトトリガーを全削除
  deleteTrigger();
  //トリガー用変数の初期化
  Properties.setProperty("task_page",'');
}