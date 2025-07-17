// ===============================================================
// 各種設定
// ===============================================================
const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  
  // ⚠️🔻【要設定】ここにSlackのWebhook URLを貼り付けてください！ 🔻⚠️
  SLACK_WEBHOOK_URL: "",

  // 「WBS」シートに関する設定
  WBS: {
    SHEET_NAME: 'WBS',
    DATA_START_ROW: 2,
    COLUMNS: {
      NO: 0, TASK_NAME: 3, START_DATE: 6, END_DATE: 7, STATUS: 8
    },
    STATUSES: {
      // WBSシート - 終了日超過タスクで除外するステータス
      EXCLUDED_FOR_OVERDUE: ['完了', '対応不要'],

      // WBSシート - 今週開始タスクで対象とするステータス
      TARGET_FOR_THIS_WEEK: ['未着手', '保留']
    }
  },

  // 「TODO管理」シートに関する設定
  TODO: {
    SHEET_NAME: 'TODO管理',
    DATA_START_ROW: 3,
    COLUMNS: {
      NO: 0, ENTRY_DATE: 4, ASSIGNEE: 5, CONTENT: 6, DUE_DATE: 7, STATUS: 8
    },
    STATUSES: {
      // TODO管理シート - 通知対象から除外するステータス（「完了」のみ）
      EXCLUDED: ['完了']
    }
  },

  // 「課題管理」シートに関する設定
  ISSUE: {
    SHEET_NAME: '課題管理',
    DATA_START_ROW: 3,
    COLUMNS: {
      NO: 0, ISSUE_CONTENT: 6, RESPONSE_CONTENT: 7, DUE_DATE: 8, STATUS: 9
    },
    STATUSES: {
      // 課題管理シート - 通知対象から除外するステータス（「完了」のみ）
      EXCLUDED: ['完了']
    }
  }
};

/**
 * メインの処理を実行
 */
function checkTasksAndNotifySlack() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const wbsSheet = spreadsheet.getSheetByName(CONFIG.WBS.SHEET_NAME);
    const todoSheet = spreadsheet.getSheetByName(CONFIG.TODO.SHEET_NAME);
    const issueSheet = spreadsheet.getSheetByName(CONFIG.ISSUE.SHEET_NAME);

    if (!wbsSheet) {
      Logger.log(`エラー: WBSシート '${CONFIG.WBS.SHEET_NAME}' が見つかりません。処理をスキップ。`);
    } else {
      processWBSSheet(wbsSheet);
    }

    if (!todoSheet) {
      Logger.log(`エラー: TODO管理シート '${CONFIG.TODO.SHEET_NAME}' が見つかりません。処理をスキップ。`);
    } else {
      processTodoSheet(todoSheet);
    }

    if (!issueSheet) {
      Logger.log(`エラー: 課題管理シート '${CONFIG.ISSUE.SHEET_NAME}' が見つかりません。処理をスキップ。`);
    } else {
      processIssueSheet(issueSheet);
    }

  } catch (e) {
    Logger.log('予期せぬエラーが発生しました: ' + e.toString());
  }
}

/**
 * WBSシートの処理を実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function processWBSSheet(sheet) {
  Logger.log(`--- ${sheet.getName()} シートの処理を開始 ---`);
  
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit`;
  const sheetId = sheet.getSheetId();

  const today = new Date(sheet.getRange('I1').getValue()); 
  today.setHours(0, 0, 0, 0);

  const startOfThisWeek = getStartOfWeek(today);
  const endOfThisWeek = new Date(startOfThisWeek);
  endOfThisWeek.setDate(startOfThisWeek.getDate() + 4);
  endOfThisWeek.setHours(23, 59, 59, 999);

  const allData = sheet.getDataRange().getValues().slice(CONFIG.WBS.DATA_START_ROW - 1);
  const overdueTasks = [];
  const thisWeekStartTasks = [];

  allData.forEach((row, index) => {
    const rowIndex = index + CONFIG.WBS.DATA_START_ROW;
    const taskName = row[CONFIG.WBS.COLUMNS.TASK_NAME];
    if (!taskName) return;

    const link = `<${spreadsheetUrl}#gid=${sheetId}&range=A${rowIndex}|行:${rowIndex}>`;

    const taskNo = row[CONFIG.WBS.COLUMNS.NO] || 'なし';
    const status = String(row[CONFIG.WBS.COLUMNS.STATUS] || "").trim();
    const startDate = parseSheetDate(row[CONFIG.WBS.COLUMNS.START_DATE]);
    const endDate = parseSheetDate(row[CONFIG.WBS.COLUMNS.END_DATE]);

    if (endDate && endDate.getTime() < today.getTime()) {
      if (!CONFIG.WBS.STATUSES.EXCLUDED_FOR_OVERDUE.includes(status)) {
        const formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        overdueTasks.push(`・${link} No${taskNo} ${taskName} (終了日: ${formattedEndDate}) [ステータス: ${status}]`);
      }
    }

    if (startDate && startDate.getTime() >= startOfThisWeek.getTime() && startDate.getTime() <= endOfThisWeek.getTime()) {
      if (CONFIG.WBS.STATUSES.TARGET_FOR_THIS_WEEK.includes(status)) {
        const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        thisWeekStartTasks.push(`・${link} No${taskNo} ${taskName} (開始日: ${formattedStartDate}) [ステータス: ${status}]`);
      }
    }
  });

  if (overdueTasks.length > 0) {
    const intro = `以下のタスクの終了日が ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/M/d')} より前です。`;
    const messageForLog = `*【WBSシート - 注意】終了日が過ぎたタスクがあります！*\n\n${intro}\n\n${overdueTasks.join('\n').replace(/<[^>]+>/g, '')}`;
    Logger.log('--- WBSシート: 終了日が過ぎたタスク ---');
    Logger.log(messageForLog);
    const messageForSlack = createSlackNotificationMessage("*【WBSシート - 注意】終了日が過ぎたタスクがあります！*", intro, overdueTasks);
    sendToSlack(messageForSlack);
  } else {
    Logger.log('WBSシート: 終了日が過ぎたタスクはありませんでした。');
    sendToSlack('【WBSシート】終了日が過ぎたタスクはありませんでした。素晴らしい！');
  }

  if (thisWeekStartTasks.length > 0) {
    const intro = `以下のタスクが今週 (${Utilities.formatDate(startOfThisWeek, Session.getScriptTimeZone(), 'yyyy/M/d')} - ${Utilities.formatDate(endOfThisWeek, Session.getScriptTimeZone(), 'yyyy/M/d')}) 開始予定です。`;
    const messageForLog = `*【WBSシート - お知らせ】今週開始のタスクです！*\n\n${intro}\n\n${thisWeekStartTasks.join('\n').replace(/<[^>]+>/g, '')}`;
    Logger.log('--- WBSシート: 今週開始タスク ---');
    Logger.log(messageForLog);
    const messageForSlack = createSlackNotificationMessage("*【WBSシート - お知らせ】今週開始のタスクです！*", intro, thisWeekStartTasks);
    sendToSlack(messageForSlack);
  } else {
    Logger.log('WBSシート: 今週開始のタスクはありませんでした。');
  }

  Logger.log(`--- ${sheet.getName()} シートの処理が完了しました ---`);
}

/**
 * TODO管理シートの処理を実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function processTodoSheet(sheet) {
  Logger.log(`--- ${sheet.getName()} シートの処理を開始 ---`);

  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit`;
  const sheetId = sheet.getSheetId();
  
  const allData = sheet.getDataRange().getValues();
  if (allData.length < CONFIG.TODO.DATA_START_ROW) {
    Logger.log('TODO管理シート: 抽出するTODOが見つかりませんでした (データ行が少ない可能性があります)。');
    return;
  }

  const extractedTasks = [];
  for (let i = CONFIG.TODO.DATA_START_ROW - 1; i < allData.length; i++) {
    const row = allData[i];
    const entryValue = row[CONFIG.TODO.COLUMNS.ENTRY_DATE];
    const status = String(row[CONFIG.TODO.COLUMNS.STATUS] || "").trim();

    if (entryValue !== "" && !CONFIG.TODO.STATUSES.EXCLUDED.includes(status)) {
      extractedTasks.push({ row: row, rowIndex: i + 1 });
    }
  }

  if (extractedTasks.length === 0) {
    Logger.log('TODO管理シート: 未完了のTODOはありませんでした。素晴らしい！');
    return;
  }

  // ログ出力用のメッセージ（テキストリンク無し）
  const todoMessagesForLog = extractedTasks.map(item => {
    const row = item.row;
    const todoNo = row[CONFIG.TODO.COLUMNS.NO] || 'なし';
    const assignee = row[CONFIG.TODO.COLUMNS.ASSIGNEE] || 'なし';
    const content = row[CONFIG.TODO.COLUMNS.CONTENT] || 'なし';
    const dueDate = row[CONFIG.TODO.COLUMNS.DUE_DATE] ? Utilities.formatDate(parseSheetDate(row[CONFIG.TODO.COLUMNS.DUE_DATE]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'なし';
    const status = String(row[CONFIG.TODO.COLUMNS.STATUS] || "").trim();
    return `・No.${todoNo} 対応者: ${assignee} 期日: ${dueDate} [ステータス: ${status}] \n 内容: ${content}`;
  }).join('\n\n');

  const messageForLog = `*【TODO管理シート - 注意】未完了のタスクがあります！*\n\n${todoMessagesForLog}`;
  Logger.log('--- TODO管理シート: 抽出された未完了タスク ---');
  Logger.log(messageForLog);

  // Slack通知用のメッセージ（テキストリンク付きにする）
  const todoMessagesForSlack = extractedTasks.map(item => {
    const rowData = item.row;
    const rowIndex = item.rowIndex;

    const link = `<${spreadsheetUrl}#gid=${sheetId}&range=A${rowIndex}|行:${rowIndex}>`;
    
    const todoNo = rowData[CONFIG.TODO.COLUMNS.NO] || 'なし';
    const assignee = rowData[CONFIG.TODO.COLUMNS.ASSIGNEE] || 'なし';
    const content = rowData[CONFIG.TODO.COLUMNS.CONTENT] || 'なし';
    const dueDate = rowData[CONFIG.TODO.COLUMNS.DUE_DATE] ? Utilities.formatDate(parseSheetDate(rowData[CONFIG.TODO.COLUMNS.DUE_DATE]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'なし';
    const status = String(rowData[CONFIG.TODO.COLUMNS.STATUS] || "").trim();
    
    return `${link} 【No.${todoNo} 対応者: ${assignee} 期日: ${dueDate} [ステータス: ${status}]】 \n 内容: ${content}`;
  }).join('\n◼︎━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◼︎\n');

  const messageForSlack = createSlackNotificationMessage(
    "*【TODO管理シート - 注意】未完了のタスクがあります！*",
    "", 
    [todoMessagesForSlack]
  );
  sendToSlack(messageForSlack);

  Logger.log(`--- ${sheet.getName()} シートの処理が完了しました ---`);
}

/**
 * 課題管理シートの処理を実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function processIssueSheet(sheet) {
  Logger.log(`--- ${sheet.getName()} シートの処理を開始 ---`);

  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit`;
  const sheetId = sheet.getSheetId();
  
  const allData = sheet.getDataRange().getValues();
  if (allData.length < CONFIG.ISSUE.DATA_START_ROW) {
    Logger.log('課題管理シート: 抽出する課題が見つかりませんでした (データ行が少ない可能性があります)。');
    return;
  }

  const extractedIssues = [];
  for (let i = CONFIG.ISSUE.DATA_START_ROW - 1; i < allData.length; i++) {
    const row = allData[i];
    const issueContent = row[CONFIG.ISSUE.COLUMNS.ISSUE_CONTENT];
    const status = String(row[CONFIG.ISSUE.COLUMNS.STATUS] || "").trim();

    if (issueContent !== "" && !CONFIG.ISSUE.STATUSES.EXCLUDED.includes(status)) {
      extractedIssues.push({ row: row, rowIndex: i + 1 });
    }
  }

  if (extractedIssues.length === 0) {
    Logger.log('課題管理シート: 未完了の課題はありませんでした。');
    return;
  }

  const issueMessagesForLog = extractedIssues.map(item => {
    const row = item.row;
    const issueNo = row[CONFIG.ISSUE.COLUMNS.NO] || 'なし';
    const issueContent = row[CONFIG.ISSUE.COLUMNS.ISSUE_CONTENT] || 'なし';
    const responseContent = row[CONFIG.ISSUE.COLUMNS.RESPONSE_CONTENT] || 'なし';
    const dueDate = row[CONFIG.ISSUE.COLUMNS.DUE_DATE] ? Utilities.formatDate(parseSheetDate(row[CONFIG.ISSUE.COLUMNS.DUE_DATE]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'なし';
    const status = String(row[CONFIG.ISSUE.COLUMNS.STATUS] || "").trim();
    return `・No.${issueNo} 期日: ${dueDate} [結果: ${status}] \n 課題内容: ${issueContent}\n 対応内容: ${responseContent}`;
  }).join('\n\n');

  const messageForLog = `*【課題管理シート - 注意】未完了の課題があります！*\n\n${issueMessagesForLog}`;
  Logger.log('--- 課題管理シート: 抽出された未完了課題 ---');
  Logger.log(messageForLog);

  const issueMessagesForSlack = extractedIssues.map(item => {
    const rowData = item.row;
    const rowIndex = item.rowIndex;

    const link = `<${spreadsheetUrl}#gid=${sheetId}&range=A${rowIndex}|行:${rowIndex}>`;
    
    const issueNo = rowData[CONFIG.ISSUE.COLUMNS.NO] || 'なし';
    const issueContent = rowData[CONFIG.ISSUE.COLUMNS.ISSUE_CONTENT] || 'なし';
    const responseContent = rowData[CONFIG.ISSUE.COLUMNS.RESPONSE_CONTENT] || 'なし';
    const dueDate = rowData[CONFIG.ISSUE.COLUMNS.DUE_DATE] ? Utilities.formatDate(parseSheetDate(rowData[CONFIG.ISSUE.COLUMNS.DUE_DATE]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'なし';
    const status = String(rowData[CONFIG.ISSUE.COLUMNS.STATUS] || "").trim();
    
    return `${link} 【No.${issueNo} 期日: ${dueDate} [結果: ${status}]】 \n 課題内容: ${issueContent}\n 対応内容: ${responseContent}`;
  }).join('\n◼︎━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◼︎\n');

  const messageForSlack = createSlackNotificationMessage(
    "*【課題管理シート - 注意】未完了の課題があります！*",
    "", 
    [issueMessagesForSlack]
  );
  sendToSlack(messageForSlack);

  Logger.log(`--- ${sheet.getName()} シートの処理が完了しました ---`);
}


// ==================================================================
// ヘルパー関数
// ==================================================================

function createSlackNotificationMessage(title, introduction, itemList) {
  let message = title;
  if (introduction) { message += `\n\n${introduction}`; }
  if (itemList && itemList.length > 0) { message += `\n\n${itemList.join('\n')}`; }
  return message;
}

function parseSheetDate(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  try {
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  } catch (e) {}
  return null;
}

function getStartOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function sendToSlack(message) {
  if (!CONFIG.SLACK_WEBHOOK_URL) {
    Logger.log('エラー: Slack Webhook URLが設定されていません。スクリプトプロパティ「SLACK_WEBHOOK_URL」を確認してください。');
    return;
  }
  const payload = { "text": message, "mrkdwn": true }; // mrkdwnを有効にする
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, options);
    Logger.log('Slackへの通知に成功しました。');
  } catch (e) {
    Logger.log('Slackへの通知に失敗しました: ' + e.toString());
  }
}
