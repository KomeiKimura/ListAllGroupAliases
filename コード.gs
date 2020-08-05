/**
 * List up all groups and group aliases
 */
function listAllGroupsAliases()
{  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  var pageToken, page;
  var rows = [];
  
  // Push cloumn name
  rows.push(['Group Name','Group Email','Group Alias']);
  
  do {
    // Get groups
    page = AdminDirectory.Groups.list({
      domain: 'uluru.jp',
      maxResults: 10000,
      pageToken: pageToken
    });    
    var groups = page.groups;
    
    if (groups) {
      for (var i = 0; i < groups.length; i++) {
        
        // Get members
        var group = groups[i];
        var aliases = page.aliases;
        if (aliases){
          for (var l = 0;l < aliases.length;l++){
            
            // Create cols => Push row
            var member = members[l];
            var cols = [];
            cols.push(group.name);
            cols.push(group.email);
            cols.push(group.aliases);
            rows.push(cols);
          }
        }
      }
      
      // sheet へ書き込み
      sheet.getRange(1, 1, rows.length, 3).setValues(rows);
      
    } else {
      Logger.log('グループがありません');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}

/**
 * On open function
 * - Create menu to get list.
 */
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu("管理")
  .addItem("グループエイリアス一覧取得", "listAllGroupsAliases")
  .addToUi();
}

