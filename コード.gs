var form_id = '1zpr5c46YWqJqhTBMIFAr9X8bXezJtOdP6OSkoLptvLs'; // フォームID
var question_title = 'ご希望の日程をお選びください'; // 動的にしたい質問のタイトル
var sheet_id = '1Psa8ngX9_Q5sP4YS3P7DR0ACQoHK_tiSh9D9uZaK0qE'; // スプレッドシートID
var sheet_name = 'シート1'; // シート名

// スプレッドシートの内容でフォームの選択項目を更新する
function change_pulldown_list(){

  // Googleフォームから質問データを取得
  var form = FormApp.openById(form_id);
  var target_item;
  var items = form.getItems();
  for(var i = 0; i < items.length; i++){
    if (items[i].getTitle() === question_title) {
      target_item = items[i];
    }
  }
  if (!target_item) {
    // 見つからなければ終了
    return;
  }

  // スプレッドシートからデータ取得
  var ss = SpreadsheetApp.openById(sheet_id);
  var sheet = ss.getSheetByName(sheet_name);
  var last_row = sheet.getLastRow();
  // ※A列固定
  var question_list = sheet.getRange(1, 1, last_row, 1).getDisplayValues();

  // 質問の選択項目を更新
  target_item.asListItem().setChoiceValues(question_list);
}

// フォームがsubmitされたときに回答された選択項目をフォームから削除する
function trigger_submit(event){
  var itemResponses = event.response.getItemResponses();

  // 回答を取得
  var target_value = '';
  for (var i = 0; i < itemResponses.length; i++) {
    if (itemResponses[i].getItem().getTitle() === question_title){
      var answer = itemResponses[i].getResponse();
      target_value = answer;
    }
  }
  
  // 削除したくない選択肢の場合は何もしない
  if (target_value === '上記以外の日程で相談したい') {
     return;
  }
  
  if (target_value) {
    // スプレッドシートから削除して
    delete_row(target_value);
    // フォームに反映
    change_pulldown_list();
  }
}

// スプレッドシートから引数で渡された値を持つ行を削除する
function delete_row(target_value){
  var ss = SpreadsheetApp.openById(sheet_id);
  var sheet = ss.getSheetByName(sheet_name);
  var last_row = sheet.getLastRow();
  for(var i = 1; i <= last_row; i++){
    // ※A列固定
    var range = ss.getRange('A' + i);
    var value = range.getDisplayValue();
    if(value == target_value){
      var start_row = i;
      var num_row = 1;
      ss.deleteRows(start_row, num_row);
      i = i - 1;
    }
  }
}