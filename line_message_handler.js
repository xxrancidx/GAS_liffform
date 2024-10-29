// LINEのメッセージ送受信設定
const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
const LINE_URL = 'https://api.line.me/v2/bot/message/reply';

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ 'content': 'get ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    // 応答用Tokenを取得
    const event = JSON.parse(e.postData.contents).events[0];
    if (!event || !event.replyToken) {
      throw new Error('Invalid event data');
    }
    
    const replyToken = event.replyToken;
    const userMessage = event.message.text;
    
    if (!userMessage) {
      throw new Error('No message content');
    }

    // メッセージをパースして構造化データに変換
    const parsedData = parseMessage(userMessage);
    
    // スプレッドシートにデータを保存
    saveToSpreadsheet(parsedData);

    // 返信メッセージを送信
    sendLineResponse(replyToken);

    return ContentService.createTextOutput(
      JSON.stringify({ 'content': 'post ok' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService.createTextOutput(
      JSON.stringify({ 'content': 'error occurred', 'error': error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function parseMessage(message) {
  try {
    const lines = message.split('\n');
    const data = {
      identifier: '',
      date: null,
      name: '',
      assignmentType: '',
      siteName: '',
      hasOtherReports: '',
      otherReports: []
    };

    lines.forEach(line => {
      if (line.startsWith('識別子：')) data.identifier = line.replace('識別子：', '');
      else if (line.startsWith('出向日：')) {
        const dateStr = line.replace('出向日：', '').trim();
        data.date = dateStr;
      }
      else if (line.startsWith('氏名：')) data.name = line.replace('氏名：', '');
      else if (line.startsWith('出向内容：')) data.assignmentType = line.replace('出向内容：', '');
      else if (line.startsWith('現場名：')) data.siteName = line.replace('現場名：', '');
      else if (line.startsWith('他の出向報告：')) data.hasOtherReports = line.replace('他の出向報告：', '');
      else if (line.match(/^\d+\./)) {
        const reportContent = line.replace(/^\d+\.\s*/, '');
        data.otherReports.push(parseReportContent(reportContent));
      }
    });

    return data;
  } catch (error) {
    throw new Error(`Failed to parse message: ${error.message}`);
  }
}

function parseReportContent(content) {
  // 報告内容を分割（団体名/人数/現場名）
  const parts = content.split('/');
  if (parts.length !== 3) {
    throw new Error('Invalid report format');
  }

  const [organization, countInfo, siteName] = parts;
  
  // 人数情報を解析
  const counts = {
    fullDay: 0,
    halfDay: 0,
    night: 0
  };

  countInfo.split(',').forEach(count => {
    if (count.includes('全日')) {
      counts.fullDay = parseInt(count.match(/\d+/)[0] || 0);
    } else if (count.includes('半日')) {
      counts.halfDay = parseInt(count.match(/\d+/)[0] || 0);
    } else if (count.includes('夜間')) {
      counts.night = parseInt(count.match(/\d+/)[0] || 0);
    }
  });

  return {
    organization: organization.trim(),
    ...counts,
    siteName: siteName.trim()
  };
}

function saveToSpreadsheet(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = sheet.getSheetByName("data");
    const reportSheet = sheet.getSheetByName("data_report");
    
    if (!mainSheet || !reportSheet) {
      throw new Error('Required sheets not found');
    }

    // メインデータの保存
    const receiveTime = new Date();
    const mainRow = [
      receiveTime,
      data.identifier,
      data.date,
      data.name,
      data.assignmentType,
      data.siteName
    ];

    const mainLastRow = mainSheet.getLastRow() + 1;
    mainSheet.getRange(mainLastRow, 1, 1, mainRow.length).setValues([mainRow]);
    
    // 日付フォーマットの設定
    mainSheet.getRange(mainLastRow, 1).setNumberFormat('yyyy/MM/dd H:mm:ss');
    mainSheet.getRange(mainLastRow, 3).setNumberFormat('yyyy/MM/dd');

    // 追加報告の保存
    if (data.hasOtherReports === 'はい' && data.otherReports.length > 0) {
      data.otherReports.forEach(report => {
        const reportRow = [
          receiveTime,
          data.identifier,
          data.date,
          report.organization,
          report.fullDay,
          report.halfDay,
          report.night,
          report.siteName
        ];

        const reportLastRow = reportSheet.getLastRow() + 1;
        reportSheet.getRange(reportLastRow, 1, 1, reportRow.length).setValues([reportRow]);
        
        // 日付フォーマットの設定
        reportSheet.getRange(reportLastRow, 1).setNumberFormat('yyyy/MM/dd H:mm:ss');
        reportSheet.getRange(reportLastRow, 3).setNumberFormat('yyyy/MM/dd');
      });
    }
  } catch (error) {
    throw new Error(`Failed to save to spreadsheet: ${error.message}`);
  }
}

function sendLineResponse(replyToken) {
  try {
    const messages = [
      {
        'type': 'text',
        'text': "入力を受け付けました！\n本日もお疲れ様でした。😌",
      }
    ];

    const response = UrlFetchApp.fetch(LINE_URL, {
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': `Bearer ${LINE_TOKEN}`,
      },
      'method': 'post',
      'payload': JSON.stringify({
        'replyToken': replyToken,
        'messages': messages,
      }),
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`LINE API returned status ${response.getResponseCode()}`);
    }
  } catch (error) {
    throw new Error(`Failed to send LINE response: ${error.message}`);
  }
}
