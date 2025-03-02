function collectCheckpointResults() {
  var spreadsheet = SpreadsheetApp.openById("集計_SPREADSHEET_ID"); // 集計スプレッドシートのID
  var logSheet = spreadsheet.getSheetByName("実行履歴") || spreadsheet.insertSheet("実行履歴"); // 実行履歴シート

  // ✅ 名簿スプレッドシートの取得
  var teamSpreadsheet = SpreadsheetApp.openById("名簿_SPREADSHEET_ID"); // 名簿スプレッドシートのID
  var teamListSheet = teamSpreadsheet.getSheetByName("名簿DEMO");

  if (!teamListSheet) {
    Logger.log("🚨 エラー: 「名簿DEMO」シートが見つかりません！");
    return;
  }

  if (teamListSheet.getLastRow() < 2) {
    Logger.log("⚠️ 名簿DEMO シートが空です。処理をスキップします。");
    return;
  }

  var teamList = teamListSheet.getDataRange().getValues();
  var teamMap = {};  // メールアドレス → 班番号のマップ

  for (var i = 1; i < teamList.length; i++) {
    var email = teamList[i][1];  // 2列目（メールアドレス）
    var teamNumber = parseInt(teamList[i][0]); // 1列目（班番号を数値に変換）
    teamMap[email] = teamNumber;
  }

  // ✅ 各チェックポイントのスプレッドシートIDとシート名
  var formSheets = [
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "学内散策の準備" }, 
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_1" }, 
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_2" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_3" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_4" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_5" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_6" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_7" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_8" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_9" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_10" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_11" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_12" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_13" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_14" },
    { id: "チェックポイントGoogleForm由来のSPREADSHEET_ID", name: "Quiz_15" }
  ];

  formSheets.forEach(sheetInfo => {
    var formSheetId = sheetInfo.id;
    var sheetName = sheetInfo.name;

    var formSheet = SpreadsheetApp.openById(formSheetId).getSheets()[0];
    var responses = formSheet.getDataRange().getValues();

    var teamFirstResponses = {};  // 各班の最初の回答
    var teamScores = {};  // 各班の合計スコア
    var teamFirstEmail = {}; // 各班の「最も早く回答した人のメールアドレス」

    for (var j = 1; j < responses.length; j++) {
      var timestamp = responses[j][0];
      var email = responses[j][1];
      var score = parseInt(responses[j][2]) || 0; // スコアが undefined の場合 0 にする

      var teamNumber = teamMap[email];
      if (!teamNumber) continue;

      if (!teamFirstResponses[teamNumber]) {
        teamFirstResponses[teamNumber] = {};
      }

      if (!teamFirstResponses[teamNumber][formSheetId] || 
          teamFirstResponses[teamNumber][formSheetId].timestamp > timestamp) {
        teamFirstResponses[teamNumber][formSheetId] = { timestamp, email, score, teamNumber };

        if (!teamFirstEmail[teamNumber] || teamFirstEmail[teamNumber].timestamp > timestamp) {
          teamFirstEmail[teamNumber] = { timestamp, email };
        }
      }
    }

    Object.keys(teamFirstResponses).forEach(team => {
      if (!teamScores[team]) teamScores[team] = 0;
      
      Object.values(teamFirstResponses[team]).forEach(response => {
        teamScores[team] += response.score;
      });
    });

    // ✅ シートを作成 or 更新
    var resultSheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
    resultSheet.clear();
    resultSheet.appendRow(["班番号", "合計スコア", "最初に回答した人のメールアドレス"]);

    Object.keys(teamScores).forEach(team => {
      var email = teamFirstEmail[team] ? teamFirstEmail[team].email : "（メールなし）";
      resultSheet.appendRow([team, teamScores[team], email]);
    });

    Logger.log("✅ 集計完了: " + sheetName);
  });

  // ✅ 実行履歴シートに「集計完了」ログを記録
  logSheet.appendRow(["集計実行", new Date().toLocaleString()]);

  Logger.log("✅ すべてのチェックポイントの集計が完了しました！");


  summarizeTotalScores();  // 合計得点を集計
}
function summarizeTotalScores() {
  var spreadsheet = SpreadsheetApp.openById("集計用SPREADSHEET_ID"); // 集計スプレッドシート
  var summarySheet = spreadsheet.getSheetByName("合計得点") || spreadsheet.insertSheet("合計得点");
  summarySheet.clear();
  summarySheet.appendRow(["班番号", "合計得点"]);

  var teamScores = {};  // 班ごとの合計得点を保存するオブジェクト

  var formSheets = [
    "学内散策の準備", "Quiz_1", "Quiz_2", "Quiz_3", "Quiz_4", "Quiz_5", "Quiz_6", "Quiz_7",
    "Quiz_8", "Quiz_9", "Quiz_10", "Quiz_11", "Quiz_12", "Quiz_13", "Quiz_14", "Quiz_15"
  ];

  formSheets.forEach(sheetName => {
    var resultSheet = spreadsheet.getSheetByName(sheetName);
    if (!resultSheet) return;  // シートがない場合はスキップ

    var data = resultSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { // 1行目（ヘッダー）をスキップ
      var teamNumber = data[i][0];
      var score = parseInt(data[i][1]) || 0;

      if (!teamScores[teamNumber]) {
        teamScores[teamNumber] = 0;
      }
      teamScores[teamNumber] += score;
    }
  });

  // 合計得点を「合計得点」シートに書き込む
  var totalScoresData = [["班番号", "合計得点"]];
  Object.keys(teamScores).forEach(team => {
    totalScoresData.push([team, teamScores[team]]);
  });

  summarySheet.getRange(1, 1, totalScoresData.length, totalScoresData[0].length).setValues(totalScoresData);
  
  Logger.log("✅ 合計得点の集計が完了しました！");
}

