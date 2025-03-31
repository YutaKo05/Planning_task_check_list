// GASのコード
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var activeCell = sheet.getActiveCell();
  var sheetName = sheet.getName();
  var row = activeCell.getRow();
  var col = activeCell.getColumn();

  // 監視するシート名のリスト
  var sheetNames = [
    'ヒーローショー（石田）',
    'ギダイコンテスト（西森）',
    '中夜祭（成瀬）',
    'ゲーム大会（酒井）',
    'カラオケ大会（黄）',
    '技大神輿(松田・黒木)',
    'クイズ大会(冷水）',
    '灯籠ワークショップ(大川)',
    'グルメGP（新美）',
    'ギネス記録（砂山）',
    'シャボン玉広場（太閤）',
    'アート展（西村）',
    '体力テスト(加藤）',
    '気配切り(岩渕)',
    '謎解き（板谷）',
    'スタンプラリー（池上）'
  ];

  // アクティブなシートが監視するシートのリストに含まれているか確認
  if (sheetNames.indexOf(sheetName) !== -1) {
    // 変更が特定のセル範囲内か確認 (例えば、4〜7行目の2列目)
    if (col == 2 && (row >= 4 && row <= 7)) {
      // 4行目のセルがYESになった場合
      if (row == 4 && activeCell.getValue() == "YES") {
        var slackText = sheetName + "の企画書が完成しました。\n" + "https://docs.google.com/spreadsheets/d/1_oLyZ5b8RrCSCY5KQEfKV1srrNyhm0LVBWZ07_CUga8/edit?gid=0#gid=0";
        sendSlack(slackText);
      }
      
      // 5行目のセルがYESになった場合
      if (row == 5 && activeCell.getValue() == "YES") {
        var slackText = sheetName + "のマニュアルが完成しました。\n" + "https://docs.google.com/spreadsheets/d/1_oLyZ5b8RrCSCY5KQEfKV1srrNyhm0LVBWZ07_CUga8/edit?gid=0#gid=0";
        sendSlack(slackText);
      }
      
      // 6行目のセルがYESになった場合
      if (row == 6 && activeCell.getValue() == "YES") {
        var slackText = sheetName + "の物品購入リストが完成しました。\n" + "https://docs.google.com/spreadsheets/d/1_oLyZ5b8RrCSCY5KQEfKV1srrNyhm0LVBWZ07_CUga8/edit?gid=0#gid=0";
        sendSlack(slackText);
      }
      
      // 7行目のセルがDONEになった場合
      if (row == 7 && activeCell.getValue() == "DONE") {
        var slackText = sheetName + "の物品購入を完了しました。\n" + "https://docs.google.com/spreadsheets/d/1_oLyZ5b8RrCSCY5KQEfKV1srrNyhm0LVBWZ07_CUga8/edit?gid=0#gid=0";
        sendSlack(slackText);
      }
    }
  }
}

function sendSlack(slackText) {
  var webHookUrl = "https://hooks.slack.com/services/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"; // Webhook URLを設定
  var jsonData = {
    'icon_emoji': ':star2:',
    "text": slackText,
    "username": "進捗確認bot"
  };
  var payload = JSON.stringify(jsonData);
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload,
  };
  UrlFetchApp.fetch(webHookUrl, options);
}
