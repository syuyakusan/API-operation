/**
 * userIDとスプレッドシートの対応の記録方法
 * 一度スクリプトプロパティに保存→検証して大丈夫そうだったらスプレッドシートのDBに記録
 * ∵スプレッドシートのみの記録だと２つ目以降の連携の際にDBから取り出すべきURLの区別がつかなくなる
 * 個人トークの対応はスプレッドシートDBに、グループトークの対応はスクリプトプロパティに記録する
 */

/**
 * リクエストを受けたときに実行する関数
 */
function doPost(e) {
  //LINE Messaging APIのチャネルアクセストークンを設定
  let LINE_API_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_API_TOKEN');
  // WebHookで取得したJSONデータをオブジェクト化し、取得
  let eventData = JSON.parse(e.postData.contents).events[0];
  
  const eventType = eventData.type;
  //取得したデータから、応答用のトークンを取得
  let replyToken = eventData.replyToken;
  

  //初回案内
  if(eventType === "follow"){
    const message = "まずはこのLINEトークと集約さんの連携をしましょう！\n↓のように「リンク」と入力した上で参加しているグループの集約さんスプレッドシートのURLを送ってください！\n\n例:\nリンク\nhttps://docs.google.com/spreadsheets/d/********";
    postToTalk(replyToken,message);
  }if(eventType === "unfollow"){
    const userId = eventData.source.userId;
    PropertiesService.getScriptProperties().deleteProperty(userId);
    deleteUserMode(userId);
  }
  else if(eventType === "join"){
    const message = "こんにちは、集約さんbotです！日程集約のお手伝いを致します！\nまずはLINEと集約さんの連携をしましょう！\n私をメンションした上で、↓のように集約さんスプレッドシートのURLを送ってください！\n\n例:\n@集約さん\n連携\nhttps://docs.google.com/spreadsheets/d/********";
    postToTalk(replyToken,message);
  }
  // メッセージイベント
  else if(eventType ==="message"){

    //取得したデータから、ユーザーが投稿したメッセージを取得
    let userMessage = eventData.message.text;

    if(eventData.source.type === "group"&&eventData.message.text.includes("@集約さん")){
      // グループトーク
      if(userMessage.includes("連携")){
      try{
        addGroupIdToScriptProperty(eventData);
        // アクセスできるかチェック
        const groupId = eventData.source.groupId;
        const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(groupId);
        if( SpreadsheetApp.openByUrl(spreadSheetUrl).getSheetByName('集約').getSheetName() == "集約"){
        }else{
          throw new Error("URLが正しく読み取れません");
        }
      }catch(ERROR){
        postToTalk(replyToken,"連携に失敗しました\n"+ERROR);
        return
      }
      postToTalk(replyToken,"連携に成功しました！\n次に「@集約さん」とだけ打ってみましょう！");
      }


      // 集約モード
      if(userMessage.includes("日程集約")){
        const groupId = eventData.source.groupId;
        const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(groupId);
        const spreadSheet = SpreadsheetApp.openByUrl(spreadSheetUrl);
        // 集約する
          checkDiffForApi(spreadSheet);
          setSharedResultTextForApi(spreadSheet);

          const setSheet = spreadSheet.getSheetByName("ホーム");
          const result = setSheet.getRange("F18").getValue();
          const replyMessage = "集約が完了しました！\n"+result;

          postToTalk(replyToken,replyMessage);
      }else if(userMessage.includes("回答状況")){
        try{
          const groupId = eventData.source.groupId;
          const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(groupId);
          const spreadSheet = SpreadsheetApp.openByUrl(spreadSheetUrl);
          // 入力状況を返す
            summarizeFormsForApi(spreadSheet);
            setAnswerStatusForApi(spreadSheet);

            const setSheet = spreadSheet.getSheetByName("ホーム");
            const reqNum = setSheet.getRange("F6").getValue();
            let statusArray = setSheet.getRange(17,10,reqNum,3).getDisplayValues();
            let tmpArray = [];
            for(i=0;i<reqNum;i++){
              tmpArray.push(statusArray[i].join(" "));
            }
            const outputArray = tmpArray.join("\n");
            const replyMessage = "現在の回答状況\n"+outputArray;

            postToTalk(replyToken,replyMessage);
        }catch(ERROR){
          console.log(ERROR);
        }
      }else if(userMessage.includes("リンク")){
        const groupId = eventData.source.groupId;
        const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(groupId);
        const spreadSheet = SpreadsheetApp.openByUrl(spreadSheetUrl);
        // リンク集を返す
          const setSheet = spreadSheet.getSheetByName("ホーム");
          const replyMessage = setSheet.getRange("F17").getValue()+"\n\n集約さん公式アカウントとの個人トークからはさらに簡単に回答いただけます！";

          postToTalk(replyToken,replyMessage);
      }else{
        const replyMessage = "お困りですか？\n「@集約さん」+合言葉でお手伝い致します！\n\n↓合言葉↓\n「日程集約」：回答を集約して結果をお伝えします\n「回答状況」：メンバーの回答状況をお伝えします\n「リンク」：スプレッドシートとGoogle Fromsのリンクをお伝えします\n\nさらに詳しい情報はこちらから\nhttps://github.com/syuyakusan"

        postToTalk(replyToken,replyMessage);
      }
    }
    else if(eventData.source.type === "user"){
      // 個人トーク
      // 連携モード１段階目
      if(userMessage.includes("連携")){
        const userId = eventData.source.userId;
        const urls = getSpreadSheetUrlsByUserId(userId);
        const outputText = urls.map(value => SpreadsheetApp.openByUrl(value).getName()).join("\n");

        const message = "このLINEトークと集約さんの連携をしましょう！\n↓のように「リンク」と入力した上で参加しているグループの集約さんスプレッドシートのURLを送ってください！解除する場合は「リンク」を「解除」に置き換えてURLを送ってください！\n\n例:\nリンク\nhttps://docs.google.com/spreadsheets/d/********"+`\n\nあなたの連携状況\n${outputText}`;
    postToTalk(replyToken,message);
      }
      if(userMessage.includes("リンク")){

        try{
          addUserIdToScriptProperty(eventData);
          // アクセスできるかチェック
          const userId = eventData.source.userId;
          const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(userId);
          if( SpreadsheetApp.openByUrl(spreadSheetUrl).getSheetByName('集約').getSheetName() == "集約"){
          }else{
            throw new Error("URLが正しく読み取れません");
          }

          // 名前の確認
          const spreadSheetName = SpreadsheetApp.openByUrl(spreadSheetUrl).getName();
          const memberList = getMemberListFromSpreadSheet(spreadSheetUrl);
          let outputText=[];
          for(i=0;i<memberList.length;i++){
            outputText.push(`${memberList[i]} -> ${i+1}`);
          }
          outputText = outputText.join("\n");
          const message = `「${spreadSheetName}」と連携します！\nあなたのお名前に対応する番号を半角数字で入力してください\n${outputText}`;
          memoryUserMode(userId,"userLink");
          postToTalk(replyToken,message);

        }catch(ERROR){
          postToTalk(replyToken,"連携に失敗しました\n"+ERROR);
          return
        }

      }else if(userMessage.includes("解除")){
        try{
          addUserIdToScriptProperty(eventData);
          // アクセスできるかチェック
          const userId = eventData.source.userId;
          const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(userId);
          if( SpreadsheetApp.openByUrl(spreadSheetUrl).getSheetByName('集約').getSheetName() == "集約"){
          }else{
            throw new Error("URLが正しく読み取れません");
          }

          deleteSpreadSheetUrlsForUserId(userId,spreadSheetUrl);

          // 名前の確認
          const spreadSheetName = SpreadsheetApp.openByUrl(spreadSheetUrl).getName();
          const memberList = getMemberListFromSpreadSheet(spreadSheetUrl);
          let outputText=[];
          for(i=0;i<memberList.length;i++){
            outputText.push(`${memberList[i]} -> ${i+1}`);
          }
          outputText = outputText.join("\n");
          const message = `「${spreadSheetName}」の連携を解除しました。`;
          postToTalk(replyToken,message);

        }catch(ERROR){
          postToTalk(replyToken,"解除に失敗しました\n"+ERROR);
          return
        }

      }
      // 連携モード２段階目
      else if(isNumber(userMessage)&&!userMessage.includes("-")){
        const userId = eventData.source.userId;
        const bool = checkUserMode(userId,"userLink");
        if(bool){
          try{
            deleteUserMode(userId);
            const number = Number(userMessage);
            const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(userId);
            const name = SpreadsheetApp.openByUrl(spreadSheetUrl).getSheetByName("ホーム").getRange(7+number,6).getValue();
            memoryUserIdVsName(userId,name);

          }catch(ERROR){
            postToTalk(replyToken,"連携に失敗しました\n"+ERROR);
            return;
          }
        }else{
          postToTalk(replyToken,"正しい番号を半角数字で入力してください");
          return;
        }
        postToTalk(replyToken,"連携に成功しました！\n次にスマホの方は画面下部のヘルプボタンを、それ以外の方は「ヘルプ」と入力してみましょう！");

        moveSpreadSheetUrlDataToSheetFromScriptProperty(userId);
        
      }
      // 日程回答モード1段階目
      else if(userMessage.includes("日程回答")){
        const userId = eventData.source.userId;
        memoryUserMode(userId,"answer")
        postToTalk(replyToken,"集約さんに日程回答します。\n↓のように都合の良い日程を30分単位で 月日-開始時間-終了時間 の形式で半角数字とハイフンで入力してください。\n複数ある場合は改行して入力してください。\n\n例:\n1月23日 10:00~13:00\n→ 0123-1000-1300");
      }
      // 日程回答モード2段階目
      else if(checkUserMode(eventData.source.userId,"answer")&&userMessage.includes("-")){
        const userId = eventData.source.userId;
        let name;
        let spreadSheetUrlList = [];
        let spreadSheetNameList = [];
        try{
          spreadSheetUrlList = getSpreadSheetUrlsByUserId(userId);
          if(spreadSheetUrlList === []){
            throw new Error("スプレッドシートを連携してください。");
          }
        }catch(ERROR){
          postToTalk(replyToken,"連携を完了してください\n"+ERROR);
          return;
        }
        const answer = userMessage;
        const regex = /^[0-9-]+$/;
        let enterdDateText;
        if (regex.test(answer.replace("\n",""))) {
          for(let i=0;i<spreadSheetUrlList.length;i++){
            try{
                const spreadSheetUrl = spreadSheetUrlList[i];
                const spreadSheet = SpreadsheetApp.openByUrl(spreadSheetUrl);
                spreadSheetNameList.push(spreadSheet.getName());
                name = getNameFormUrl(userId,spreadSheetUrl);
                enterdDateText = summarizeLineDateAnswerToSpreadSheetForApi(spreadSheet,name,answer);
            }catch(ERROR){
              postToTalk(replyToken,"反映に失敗しました\n"+ERROR);
              return;
            }
          }
        }else{
          postToTalk(replyToken,"使用できるのは半角数字と-のみです")
        }
          const replyMessage = enterdDateText+"\n\n反映先:\n"+spreadSheetNameList.join("\n");

          postToTalk(replyToken,replyMessage);
      }
      else{
        const replyMessage = "お困りですか？\nスマホの方は画面下部のボタンで、それ以外の方は合言葉でお手伝い致します！\n\n↓合言葉↓\n「連携」：集約さんと連携します。複数連携すると一度の日程回答ですべての集約さんに回答できます。\n「日程回答」：日程の回答をして頂けます\n「ヘルプ」：使い方を説明します\n\nさらに詳しい情報はこちらから\nhttps://github.com/syuyakusan";

        postToTalk(replyToken,replyMessage);
      }

    }
  }
    

  
}


/**
 * LINEトークに投稿する関数
 * @param {String} replyToken
 * @param {String} replyMessage
 */
function postToTalk(replyToken,replyMessage){
  //LINE Messaging APIのチャネルアクセストークンを設定
  let LINE_API_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_API_TOKEN');

  // 応答メッセージ用のAPI URLを定義
  let url = 'https://api.line.me/v2/bot/message/reply';

  //APIリクエスト時にセットするペイロード値を設定する
  let payload = {
    'replyToken': replyToken,
    'messages': [{
        'type': 'text',
        'text': replyMessage
      }]
  };
  //HTTPSのPOST時のオプションパラメータを設定する
  let options = {
    'payload' : JSON.stringify(payload),
    'myamethod'  : 'POST',
    'headers' : {"Authorization" : "Bearer " + LINE_API_TOKEN},
    'contentType' : 'application/json'
  };
  //LINE Messaging APIにリクエストし、ユーザーからの投稿に返答する
  UrlFetchApp.fetch(url, options);
}

/**
 * 集約さん上のメンバーリストを取得する関数
 * @param {String} spreadSheetUrl
 * @returns {Array} memberList
 */
function getMemberListFromSpreadSheet(spreadSheetUrl){
  const spreadSheet = SpreadsheetApp.openByUrl(spreadSheetUrl);
  const setSheet = spreadSheet.getSheetByName("ホーム");
  const reqNum = setSheet.getRange("F6").getValue();
  const memberList = setSheet.getRange(8,6,reqNum,1).getValues().flat();
  return memberList;

}

/**
 * LINEのgroupIdと集約さんのURLをスクリプトプロパティに追加する関数
 * @param {String} userMessage ユーザーが投稿したメッセージ
 * @returns {Boolean} bool エラー判定用
 */
function addGroupIdToScriptProperty(eventData){
  let bool;
  const userMessage = eventData.message.text;
  const spreadSheetUrl = userMessage.replace("@集約さん","").replace("\n","").replace("連携","");
  // TODO:無駄な文字列が含まれていたときの処理
  const groupId = eventData.source.groupId;
  PropertiesService.getScriptProperties().setProperty(groupId,spreadSheetUrl);
  
}

/**
 * LINEのuserIdと集約さんのURLをスクリプトプロパティに追加して一時保存する関数
 * @param {String} userMessage ユーザーが投稿したメッセージ
 */
function addUserIdToScriptProperty(eventData){
  const userMessage = eventData.message.text;
  const spreadSheetUrl = userMessage.replace("@集約さん","").replace("\n","").replace("リンク","").replace("解除","");
  // TODO:無駄な文字列が含まれていたときの処理
  const userId = eventData.source.userId;
  PropertiesService.getScriptProperties().setProperty(userId,spreadSheetUrl);
  
}

/**LINEのuserIdと集約さんのURLのデータをスクリプトプロパティから消しスプレッドシートに移す関数
 * @param {String} userId
 */
function moveSpreadSheetUrlDataToSheetFromScriptProperty(userId){
  const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(userId);

  const memorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-URLs");
  let userIdList =  memorySheet.getRange(1,1,memorySheet.getLastRow(),1).getValues().flat();
  let tableMatrix =  memorySheet.getRange(1,2,memorySheet.getLastRow(),memorySheet.getLastColumn()-1).getValues();

  let userIdIndex = userIdList.indexOf(userId);
  if(userIdIndex == -1){
    userIdList.push(userId);
    tableMatrix.push([""]);
    userIdIndex = userIdList.length - 1;
  }
  let userArray = tableMatrix[userIdIndex];

  let urlAmount = 1;
  if(userArray[0] !== ""){
    urlAmount = userArray[0]+1;
  }
  userArray[0] = urlAmount;
  userArray[urlAmount] = spreadSheetUrl;
  
  const row = userIdIndex + 2;

  // 行列を転置する関数
  const transposeArray = array => array[0].map((_, colIndex) => array.map(row => row[colIndex]));
  userIdList = transposeArray([userIdList]);

  memorySheet.getRange(row-1,2,1,userArray.length).setValues([userArray]);
  memorySheet.getRange(1,1,userIdList.length,1).setValues(userIdList);
  PropertiesService.getScriptProperties().deleteProperty(userId);


}

/**
 * userIdから集約さんのURLをスプレッドシートから取得する関数
 * @param {String} userId
 * @returns {Array} userSpreadSheetUrlList userが連携している集約さんURLの配列
 */
function getSpreadSheetUrlsByUserId(userId){
  const memorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-URLs");
  let userIdList =  memorySheet.getRange(1,1,memorySheet.getLastRow(),1).getValues().flat();
  let userIdIndex = userIdList.indexOf(userId);
  const row = userIdIndex + 1;
  const userArray = memorySheet.getRange(row,2,1,memorySheet.getLastColumn()).getValues().flat();
  const urlAmount = userArray[0];
  let userSpreadSheetUrlList = [];
  for(i=0;i<urlAmount;i++){
    userSpreadSheetUrlList.push(userArray[i+1]);
  }

  return userSpreadSheetUrlList;
}

/**
 * userIdに対する特定の集約さんのURLをDBから削除する関数
 * @param {String} userId
 * @param {String} spreadSheetUrl
 */
function deleteSpreadSheetUrlsForUserId(userId,spreadSheetUrl){
  const memorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-URLs");
  const urls = getSpreadSheetUrlsByUserId(userId);
  const urlIndex = urls.indexOf(spreadSheetUrl);
  const column = urlIndex + 3;

  let userIdList =  memorySheet.getRange(1,1,memorySheet.getLastRow(),1).getValues().flat();
  let userIdIndex = userIdList.indexOf(userId);
  const row = userIdIndex + 1;

  const urlAmount = memorySheet.getRange(row,2).getValue();

  memorySheet.getRange(row,column).clearContent();
  memorySheet.getRange(row,2).setValue(urlAmount -1);
}

/**
 * userIdと各集約さんでの名前の対応をスプレッドシートの記録する関数
 * @param {String} userId
 * @param {String} name
 */
function memoryUserIdVsName(userId,name){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const memorySheet = spreadSheet.getSheetByName("userID-Name");
  const spreadSheetUrlList = memorySheet.getRange(1,2,1,memorySheet.getLastColumn()-1).getValues().flat();
  let userIdList =  memorySheet.getRange(2,1,memorySheet.getLastRow()-1,1).getValues().flat();

  const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty(userId);

  let urlIndex = spreadSheetUrlList.indexOf(spreadSheetUrl);
  if(urlIndex == -1){
    spreadSheetUrlList.push(spreadSheetUrl);
    urlIndex = spreadSheetUrlList.length - 1;
  }
  const column = urlIndex + 2;

  let userIdIndex = userIdList.indexOf(userId);
  if(userIdIndex == -1){
    userIdList.push(userId);
    userIdIndex = userIdList.length - 1;
  }
  const row = userIdIndex + 2;

  // 行列を転置する関数
  const transposeArray = array => array[0].map((_, colIndex) => array.map(row => row[colIndex]));
  userIdList = transposeArray([userIdList]);

  memorySheet.getRange(1,2,1,spreadSheetUrlList.length).setValues([spreadSheetUrlList]);
  memorySheet.getRange(2,1,userIdList.length,1).setValues(userIdList)
  memorySheet.getRange(row,column).setValue(name);
}

/**
 * userIdから各集約さんでの名前を取得する関数
 * @param {String} userId
 * @param {String} spreadSheetUrl
 * @return {String} name
 */
function getNameFormUrl(userId,spreadSheetUrl){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const memorySheet = spreadSheet.getSheetByName("userID-Name");

  const spreadSheetUrlList = memorySheet.getRange(1,2,1,memorySheet.getLastColumn()-1).getValues().flat();
  const userIdList =  memorySheet.getRange(2,1,memorySheet.getLastRow()-1,1).getValues().flat();

  const column = spreadSheetUrlList.indexOf(spreadSheetUrl) + 2;
  const row = userIdList.indexOf(userId) + 2;

  const name = memorySheet.getRange(row,column).getValue();
  return name;
}

/**
 * ユーザーの処理モードのステータスとしてスクリプトプロパティを追加する関数
 * @param {String} userId
 * @param {String} mode
 */
function memoryUserMode(userId,mode){
  const key = `ST_${userId}`;
  PropertiesService.getScriptProperties().setProperty(key,mode);
}
/**
 * ユーザーの処理モードのステータスを確認する関数
 * @param {String} userId
 * @param {String} mode
 * @returns {Boolean}
 */
function checkUserMode(userId,mode){
  const key = `ST_${userId}`;
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return (value === mode);
}
/**
 * ユーザーの処理モードのステータスを削除する関数
 */
function deleteUserMode(userId){
  const key = `ST_${userId}`;
  PropertiesService.getScriptProperties().deleteProperty(key);
}