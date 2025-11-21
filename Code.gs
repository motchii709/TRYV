/**
 * 曜日ベースのイベント管理システム - GASバックエンド
 * Google Apps Script code for weekday event management
 */

// スプレッドシートのID（デプロイ時に自動取得）
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAME = 'Events';

/**
 * Webアプリとして公開する際のエントリーポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('曜日ベースのイベント管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * スプレッドシートの初期化
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // ヘッダー行を追加
    sheet.getRange(1, 1, 1, 8).setValues([[
      'ID', '曜日', '開始時刻', '終了時刻', 'タイトル', '担当者', '説明', '色'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * 時刻を HH:MM 形式の文字列に変換
 */
function formatTime(timeValue) {
  if (!timeValue) return '';
  
  // すでに文字列の場合はそのまま返す
  if (typeof timeValue === 'string') {
    return timeValue;
  }
  
  // Dateオブジェクトの場合はフォーマット
  if (timeValue instanceof Date) {
    const hours = timeValue.getHours().toString().padStart(2, '0');
    const minutes = timeValue.getMinutes().toString().padStart(2, '0');
    return `${hours}:${minutes}`;
  }
  
  return String(timeValue);
}

/**
 * 全イベントを取得
 */
function getEvents() {
  try {
    Logger.log('=== getEvents 開始 ===');
    const sheet = initializeSpreadsheet();
    const lastRow = sheet.getLastRow();
    Logger.log('lastRow: ' + lastRow);
    
    if (lastRow <= 1) {
      Logger.log('データなし - 空配列を返します');
      return [];
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    Logger.log('取得した行数: ' + data.length);
    
    const events = data
      .filter(row => {
        const hasId = row[0] && row[0] !== '';
        return hasId;
      })
      .map(row => ({
        id: String(row[0]),
        weekday: String(row[1]),
        startTime: formatTime(row[2]),
        endTime: formatTime(row[3]),
        title: String(row[4]),
        organizer: String(row[5]),
        description: String(row[6] || ''),
        color: String(row[7] || '#4285F4')
      }));
    
    Logger.log('フィルタ後のイベント数: ' + events.length);
    Logger.log('返すイベント: ' + JSON.stringify(events));
    Logger.log('=== getEvents 終了 ===');
    
    return events;
  } catch (error) {
    Logger.log('getEvents エラー: ' + error.message);
    console.error('Error in getEvents:', error);
    throw new Error('イベントの取得に失敗しました: ' + error.message);
  }
}

/**
 * 新規イベントを追加
 */
function addEvent(eventData) {
  try {
    const sheet = initializeSpreadsheet();
    const id = Utilities.getUuid();
    
    // 時刻に'を付けて文字列として保存（自動変換を防ぐ）
    sheet.appendRow([
      id,
      eventData.weekday,
      "'" + eventData.startTime,  // 文字列として強制
      "'" + eventData.endTime,    // 文字列として強制
      eventData.title,
      eventData.organizer,
      eventData.description,
      eventData.color || '#4285F4'
    ]);
    
    return {
      success: true,
      id: id,
      message: 'イベントを追加しました'
    };
  } catch (error) {
    console.error('Error in addEvent:', error);
    throw new Error('イベントの追加に失敗しました: ' + error.message);
  }
}

/**
 * イベントを更新
 */
function updateEvent(eventData) {
  try {
    const sheet = initializeSpreadsheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      throw new Error('更新するイベントが見つかりません');
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === eventData.id) {
        const rowNumber = i + 2;
        // 時刻に'を付けて文字列として保存
        sheet.getRange(rowNumber, 1, 1, 8).setValues([[
          eventData.id,
          eventData.weekday,
          "'" + eventData.startTime,  // 文字列として強制
          "'" + eventData.endTime,    // 文字列として強制
          eventData.title,
          eventData.organizer,
          eventData.description,
          eventData.color || '#4285F4'
        ]]);
        
        return {
          success: true,
          message: 'イベントを更新しました'
        };
      }
    }
    
    throw new Error('指定されたIDのイベントが見つかりません');
  } catch (error) {
    console.error('Error in updateEvent:', error);
    throw new Error('イベントの更新に失敗しました: ' + error.message);
  }
}

/**
 * イベントを削除
 */
function deleteEvent(eventId) {
  try {
    const sheet = initializeSpreadsheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      throw new Error('削除するイベントが見つかりません');
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === eventId) {
        const rowNumber = i + 2;
        sheet.deleteRow(rowNumber);
        
        return {
          success: true,
          message: 'イベントを削除しました'
        };
      }
    }
    
    throw new Error('指定されたIDのイベントが見つかりません');
  } catch (error) {
    console.error('Error in deleteEvent:', error);
    throw new Error('イベントの削除に失敗しました: ' + error.message);
  }
}

/**
 * Discord設定を取得
 */
function getSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    return {
      webhookUrl: props.getProperty('DISCORD_WEBHOOK_URL') || '',
      postMessage: props.getProperty('DISCORD_POST_MESSAGE') || '今週のスケジュール'
    };
  } catch (error) {
    Logger.log('設定取得エラー: ' + error);
    throw error;
  }
}

/**
 * Discord設定を保存
 */
function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (settings.webhookUrl) {
      props.setProperty('DISCORD_WEBHOOK_URL', settings.webhookUrl);
    }
    if (settings.postMessage) {
      props.setProperty('DISCORD_POST_MESSAGE', settings.postMessage);
    }
    return { success: true };
  } catch (error) {
    Logger.log('設定保存エラー: ' + error);
    throw error;
  }
}

/**
 * Discordに画像を投稿
 */
function postToDiscord(imageData) {
  try {
    const props = PropertiesService.getScriptProperties();
    const webhookUrl = props.getProperty('DISCORD_WEBHOOK_URL');
    const message = props.getProperty('DISCORD_POST_MESSAGE') || '今週のスケジュール';
    
    if (!webhookUrl) {
      throw new Error('Discord Webhook URLが設定されていません。設定画面で設定してください。');
    }
    
    // Base64データからバイナリに変換
    const base64Data = imageData.split(',')[1];
    const binaryData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(binaryData, 'image/png', 'schedule.png');
    
    // マルチパートフォームデータを作成
    const boundary = '----WebKitFormBoundary' + Utilities.getUuid();
    const payload = Utilities.newBlob(
      '--' + boundary + '\r\n' +
      'Content-Disposition: form-data; name="content"\r\n\r\n' +
      message + '\r\n' +
      '--' + boundary + '\r\n' +
      'Content-Disposition: form-data; name="file"; filename="schedule.png"\r\n' +
      'Content-Type: image/png\r\n\r\n'
    ).getBytes();
    
    const fileBytes = blob.getBytes();
    const endBoundary = Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes();
    
    // 全データを結合
    const fullPayload = [];
    payload.forEach(b => fullPayload.push(b));
    fileBytes.forEach(b => fullPayload.push(b));
    endBoundary.forEach(b => fullPayload.push(b));
    
    // Discord Webhookに送信
    const options = {
      method: 'post',
      contentType: 'multipart/form-data; boundary=' + boundary,
      payload: fullPayload,
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200 && responseCode !== 204) {
      throw new Error('Discord APIエラー: ' + response.getContentText());
    }
    
    Logger.log('Discord投稿成功');
    return { success: true };
    
  } catch (error) {
    Logger.log('Discord投稿エラー: ' + error);
    throw error;
  }
}

/**
 * 毎週自動的にDiscordに投稿する関数
 * トリガーで毎週実行するように設定してください
 * 
 * 設定方法:
 * 1. GASエディタで「トリガー」を開く
 * 2. 「トリガーを追加」をクリック
 * 3. 実行する関数: weeklyPostToDiscord
 * 4. イベントのソース: 時間主導型
 * 5. 時間ベースのトリガー: 週タイマー
 * 6. 曜日と時刻を選択（例: 月曜日 9:00-10:00）
 */
function weeklyPostToDiscord() {
  try {
    Logger.log('=== 週次Discord投稿開始 ===');
    
    // Discord設定を確認
    const props = PropertiesService.getScriptProperties();
    const webhookUrl = props.getProperty('DISCORD_WEBHOOK_URL');
    const message = props.getProperty('DISCORD_POST_MESSAGE') || '今週のスケジュール';
    
    if (!webhookUrl) {
      Logger.log('エラー: Discord Webhook URLが設定されていません');
      throw new Error('Discord Webhook URLが設定されていません。設定画面で設定してください。');
    }
    
    // イベントデータを取得
    const events = getEvents();
    Logger.log('取得したイベント数: ' + events.length);
    
    // HTMLテーブルを生成してDiscordに投稿
    const scheduleText = generateScheduleText(events);
    
    // Discordに投稿
    const payload = {
      content: message,
      embeds: [{
        title: '📅 週間スケジュール',
        description: scheduleText,
        color: 6750404, // #6750A4 in decimal
        timestamp: new Date().toISOString(),
        footer: {
          text: 'TRYV! by moti'
        }
      }]
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200 && responseCode !== 204) {
      throw new Error('Discord APIエラー: ' + response.getContentText());
    }
    
    Logger.log('週次Discord投稿成功');
    Logger.log('=== 週次Discord投稿終了 ===');
    
    return { success: true, message: '週次スケジュールを投稿しました' };
    
  } catch (error) {
    Logger.log('週次Discord投稿エラー: ' + error);
    throw error;
  }
}

/**
 * イベントデータからテキスト形式のスケジュールを生成
 */
function generateScheduleText(events) {
  const weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const weekdayNames = {
    'Monday': '月曜日',
    'Tuesday': '火曜日',
    'Wednesday': '水曜日',
    'Thursday': '木曜日',
    'Friday': '金曜日'
  };
  
  let scheduleText = '';
  
  weekdays.forEach(day => {
    const dayEvents = events.filter(e => e.weekday === day).sort((a, b) => {
      return a.startTime.localeCompare(b.startTime);
    });
    
    if (dayEvents.length > 0) {
      scheduleText += `\n**${weekdayNames[day]}**\n`;
      dayEvents.forEach(event => {
        scheduleText += `\`${event.startTime}～${event.endTime}\` **${event.title}** (${event.organizer})\n`;
      });
    }
  });
  
  if (scheduleText === '') {
    scheduleText = '今週のイベントはありません。';
  }
  
  return scheduleText;
}

/**
 * テスト用: 手動で週次投稿を実行
 */
function testWeeklyPost() {
  weeklyPostToDiscord();
}

