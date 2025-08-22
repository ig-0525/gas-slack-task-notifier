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
      // WBSシート - 終了日超過で通知対象とするステータス
      TARGET_FOR_OVERDUE: ['未着手', '進行中', '保留'],
      // WBSシート - 今週開始タスクで通知対象とするステータス
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
      // TODO管理シート - 終了日超過で通知対象とするステータス
      TARGET_FOR_OVERDUE: ['未着手', '進行中', '保留'],
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
      // 課題管理シート - 終了日超過で通知対象とするステータス
      TARGET_FOR_OVERDUE: ['未着手', '進行中', '継続検討'],
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

    // WBSシートの処理
    if (!wbsSheet) {
      Logger.log(`エラー: WBSシート '${CONFIG.WBS.SHEET_NAME}' が見つかりません。処理をスキップ。`);
    } else {
      processWBSSheet(wbsSheet);
    }

    // TODO管理シートの処理
    if (!todoSheet) {
      Logger.log(`エラー: TODO管理シート '${CONFIG.TODO.SHEET_NAME}' が見つかりません。処理をスキップ。`);
    } else if (todoSheet.getLastRow() < CONFIG.TODO.DATA_START_ROW) {
      Logger.log(`${CONFIG.TODO.SHEET_NAME}: データ行が存在しないため、処理をスキップします。`);
      sendToSlack(`【${CONFIG.TODO.SHEET_NAME}】期日超過のTODOはありませんでした。素晴らしい！`);
    } else {
      processTodoSheet(todoSheet);
    }

    // 課題管理シートの処理
    if (!issueSheet) {
      Logger.log(`エラー: 課題管理シート '${CONFIG.ISSUE.SHEET_NAME}' が見つかりません。処理をスキップ。`);
    } else if (issueSheet.getLastRow() < CONFIG.ISSUE.DATA_START_ROW) {
      Logger.log(`${CONFIG.ISSUE.SHEET_NAME}: データ行が存在しないため、処理をスキップします。`);
      sendToSlack(`【${CONFIG.ISSUE.SHEET_NAME}】期日超過の課題はありませんでした。素晴らしい！`);
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

  const overdueTasksByStatus = {};
  const thisWeekStartTasksByStatus = {};

  const allData = sheet.getDataRange().getValues().slice(CONFIG.WBS.DATA_START_ROW - 1);

  allData.forEach((row, index) => {
    const rowIndex = index + CONFIG.WBS.DATA_START_ROW;
    const taskName = row[CONFIG.WBS.COLUMNS.TASK_NAME];
    if (!taskName) return;

    const link = `<${spreadsheetUrl}#gid=${sheetId}&range=A${rowIndex}|行:${rowIndex}>`;
    const taskNo = row[CONFIG.WBS.COLUMNS.NO] || 'なし';
    const status = String(row[CONFIG.WBS.COLUMNS.STATUS] || "").trim();
    const startDate = parseSheetDate(row[CONFIG.WBS.COLUMNS.START_DATE]);
    const endDate = parseSheetDate(row[CONFIG.WBS.COLUMNS.END_DATE]);

    // 終了日超過タスクのチェック
    if (endDate && endDate.getTime() < today.getTime()) {
      if (CONFIG.WBS.STATUSES.TARGET_FOR_OVERDUE.includes(status)) {
        if (!overdueTasksByStatus[status]) overdueTasksByStatus[status] = [];
        const formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        overdueTasksByStatus[status].push(`・${link} No.${taskNo} ${taskName} (終了日: ${formattedEndDate})`);
      }
    }

    // 今週開始タスクのチェック
    if (startDate && startDate.getTime() >= getStartOfWeek(today).getTime() && startDate.getTime() <= getEndOfThisWeek(today).getTime()) {
      if (CONFIG.WBS.STATUSES.TARGET_FOR_THIS_WEEK.includes(status)) {
        if (!thisWeekStartTasksByStatus[status]) thisWeekStartTasksByStatus[status] = [];
        const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        thisWeekStartTasksByStatus[status].push(`・${link} No.${taskNo} ${taskName} (開始日: ${formattedStartDate})`);
      }
    }
  });

  _notifyOverdueWbsTasks(overdueTasksByStatus, today);
  _notifyThisWeekWbsTasks(thisWeekStartTasksByStatus, today);
  
  Logger.log(`--- ${sheet.getName()} シートの処理が完了しました ---`);

  // --- WBSシートの内部ヘルパー関数定義 ---
  function _notifyOverdueWbsTasks(tasks, now) {
    if (Object.keys(tasks).length === 0) {
      Logger.log('WBSシート: 終了日が過ぎたタスクはありませんでした。');
      sendToSlack('【WBSシート】終了日が過ぎたタスクはありませんでした。素晴らしい！');
      return;
    }
    const intro = `以下のタスクの終了日が ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/M/d')} より前です。`;
    const messageGroups = formatGroupedMessage(tasks, CONFIG.WBS.STATUSES.TARGET_FOR_OVERDUE);
    const messageForSlack = createSlackNotificationMessage("*【WBSシート - 注意】終了日が過ぎたタスクがあります！*", intro, messageGroups);
    sendToSlack(messageForSlack);
  }

  function _notifyThisWeekWbsTasks(tasks, now) {
    if (Object.keys(tasks).length === 0) {
      Logger.log('WBSシート: 今週開始のタスクはありませんでした。');
      return;
    }
    const startOfThisWeek = getStartOfWeek(now);
    const endOfThisWeek = getEndOfThisWeek(now);
    const intro = `以下のタスクが今週 (${Utilities.formatDate(startOfThisWeek, Session.getScriptTimeZone(), 'yyyy/M/d')} - ${Utilities.formatDate(endOfThisWeek, Session.getScriptTimeZone(), 'yyyy/M/d')}) 開始予定です。`;
    const messageGroups = formatGroupedMessage(tasks, CONFIG.WBS.STATUSES.TARGET_FOR_THIS_WEEK);
    const messageForSlack = createSlackNotificationMessage("*【WBSシート - お知らせ】今週開始のタスクです！*", intro, messageGroups);
    sendToSlack(messageForSlack);
  }
}

/**
 * TODO管理シートの処理を実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function processTodoSheet(sheet) {
  Logger.log(`--- ${sheet.getName()} シートの処理を開始 ---`);

  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit`;
  const sheetId = sheet.getSheetId();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const overdueTasksByStatus = {};

  const allData = sheet.getDataRange().getValues().slice(CONFIG.TODO.DATA_START_ROW - 1);

  allData.forEach((row, index) => {
    const status = String(row[CONFIG.TODO.COLUMNS.STATUS] || "").trim();
    const dueDate = parseSheetDate(row[CONFIG.TODO.COLUMNS.DUE_DATE]);

    // 期日超過、かつ対象ステータスのタスクを抽出
    if (dueDate && dueDate.getTime() < today.getTime() && CONFIG.TODO.STATUSES.TARGET_FOR_OVERDUE.includes(status)) {
      if (!overdueTasksByStatus[status]) overdueTasksByStatus[status] = [];
      
      const rowIndex = index + CONFIG.TODO.DATA_START_ROW;
      const link = `<${spreadsheetUrl}#gid=${sheetId}&range=A${rowIndex}|行:${rowIndex}>`;
      const todoNo = row[CONFIG.TODO.COLUMNS.NO] || 'なし';
      const assignee = row[CONFIG.TODO.COLUMNS.ASSIGNEE] || 'なし';
      const content = row[CONFIG.TODO.COLUMNS.CONTENT] || 'なし';
      const formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      overdueTasksByStatus[status].push(`・${link} 【No.${todoNo} 対応者: ${assignee} 期日: ${formattedDueDate}】 \n 内容: ${content}`);
    }
  });

  if (Object.keys(overdueTasksByStatus).length > 0) {
    const messageGroups = formatGroupedMessage(overdueTasksByStatus, CONFIG.TODO.STATUSES.TARGET_FOR_OVERDUE, '\n◼︎━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◼︎\n');
    const messageForSlack = createSlackNotificationMessage("*【TODO管理シート - 注意】期日超過のTODOがあります！*", messageGroups);
    sendToSlack(messageForSlack);
  } else {
    Logger.log('TODO管理シート: 期日超過のTODOはありませんでした。');
    sendToSlack('【TODO管理シート】期日超過のTODOはありませんでした。素晴らしい！');
  }

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
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const overdueIssuesByStatus = {};

  const allData = sheet.getDataRange().getValues().slice(CONFIG.ISSUE.DATA_START_ROW - 1);

  allData.forEach((row, index) => {
    const status = String(row[CONFIG.ISSUE.COLUMNS.STATUS] || "").trim();
    const dueDate = parseSheetDate(row[CONFIG.ISSUE.COLUMNS.DUE_DATE]);

    // 期日超過、かつ対象ステータスの課題を抽出
    if (dueDate && dueDate.getTime() < today.getTime() && CONFIG.ISSUE.STATUSES.TARGET_FOR_OVERDUE.includes(status)) {
      if (!overdueIssuesByStatus[status]) overdueIssuesByStatus[status] = [];
      
      const rowIndex = index + CONFIG.ISSUE.DATA_START_ROW;
      const link = `<${spreadsheetUrl}#gid=${sheetId}&range=A${rowIndex}|行:${rowIndex}>`;
      const issueNo = row[CONFIG.ISSUE.COLUMNS.NO] || 'なし';
      const issueContent = row[CONFIG.ISSUE.COLUMNS.ISSUE_CONTENT] || 'なし';
      const responseContent = row[CONFIG.ISSUE.COLUMNS.RESPONSE_CONTENT] || 'なし';
      const formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      overdueIssuesByStatus[status].push(`・${link} 【No.${issueNo} 期日: ${formattedDueDate}】 \n 課題内容: ${issueContent}\n 対応内容: ${responseContent}`);
    }
  });

  if (Object.keys(overdueIssuesByStatus).length > 0) {
    const messageGroups = formatGroupedMessage(overdueIssuesByStatus, CONFIG.ISSUE.STATUSES.TARGET_FOR_OVERDUE, '\n◼︎━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◼︎\n');
    const messageForSlack = createSlackNotificationMessage("*【課題管理シート - 注意】期日超過の課題があります！*", messageGroups);
    sendToSlack(messageForSlack);
  } else {
    Logger.log('課題管理シート: 期日超過の課題はありませんでした。');
    sendToSlack('【課題管理シート】期日超過の課題はありませんでした。素晴らしい！');
  }

  Logger.log(`--- ${sheet.getName()} シートの処理が完了しました ---`);
}


// ==================================================================
// ヘルパー関数
// ==================================================================

function formatGroupedMessage(tasksByStatus, statusOrder, separator = '\n') {
  const messageGroups = [];
  for (const status of statusOrder) {
    if (tasksByStatus[status] && tasksByStatus[status].length > 0) {
      const tasksString = tasksByStatus[status].join(separator);
      messageGroups.push(`\n▼ ${status}\n${tasksString}`);
    }
  }
  return messageGroups;
}

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
  // 月曜始まりに調整 (day=0(日)の場合は-6、day=1(月)の場合は0、...day=6(土)の場合は-5)
  const diff = d.getDate() - day + (day === 0 ? -6 : 1); 
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getEndOfThisWeek(date) {
  const startOfThisWeek = getStartOfWeek(date);
  const endOfThisWeek = new Date(startOfThisWeek);
  endOfThisWeek.setDate(startOfThisWeek.getDate() + 4); // 月曜日から金曜日まで
  endOfThisWeek.setHours(23, 59, 59, 999);
  return endOfThisWeek;
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