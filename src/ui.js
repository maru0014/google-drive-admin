function btn_search() {
  const dialog = Browser.msgBox("検索を実行しますか？", Browser.Buttons.YES_NO);
  if (dialog === "no") {
    Browser.msgBox("処理を中断します");
    return;
  }

  const getPermissions = sheets.search.getRange("検索_権限出力").getValue();
  const q = sheets.search.getRange("検索_クエリ").getValue();
  const getSubfolders = q ? false : sheets.search.getRange("検索_サブフォルダも検索").getValue();
  const folderId = q ? "" : sheets.search.getRange("検索_対象フォルダID").getValue();

  sheets.search.getRange("検索_ステータス").setValue(`Search ... ${folderId}`);

  const resumeData = searchTask.loadData();
  const isResume = resumeData.length && Browser.msgBox("検索途中のデータが見つかりました。再開しますか？", Browser.Buttons.YES_NO);
  if (isResume === "yes") {
    // resumeDataがあれば再開する
    resumeSearch(resumeData, q, getPermissions, getSubfolders);
  } else {
    searchTask.clearData();
    search(folderId, q, getPermissions, getSubfolders);
  }

  if (searchTask.isCompleted()) {
    sheets.search.getRange("検索_ステータス").setValue(`Done`);
    sheets.search.sort(1, true);
  }
}

function btn_clearSearchResults() {
  const dialog = Browser.msgBox("検索結果をクリアしますか？", Browser.Buttons.YES_NO);
  if (dialog === "yes") {
    const lastCol = sheets.search.getLastColumn();
    const lastRow = sheets.search.getLastRow();
    sheets.search.getRange(7, 1, lastRow, lastCol).clearContent();
  }
}

function btn_removePermissions() {
  const dialog = Browser.msgBox("権限削除を実行しますか？", Browser.Buttons.OK_CANCEL);

  if (dialog === "cancel") {
    Browser.msgBox("処理を中断します");
    return;
  }

  const email = sheets.removePermission.getRange("権限削除_対象ユーザーメールアドレス").getValue();
  if (!email) {
    Browser.msgBox("対象ユーザーメールアドレスが入力されていないため処理を中断します");
    return;
  }

  const role = sheets.removePermission.getRange("権限削除_対象権限レベル").getValue();
  if (!role) {
    Browser.msgBox("対象権限レベルが入力されていないため処理を中断します");
    return;
  }

  removePermissions(email, role);
}

function btn_insertPermissions() {
  const dialog = Browser.msgBox("権限追加を実行しますか？", Browser.Buttons.OK_CANCEL);

  if (dialog === "cancel") {
    Browser.msgBox("処理を中断します");
    return;
  }

  const email = sheets.insertPermission.getRange("権限追加_対象ユーザーメールアドレス").getValue();
  if (!email) {
    Browser.msgBox("対象ユーザーメールアドレスが入力されていないため処理を中断します");
    return;
  }

  const role = sheets.insertPermission.getRange("権限追加_対象権限レベル").getValue();
  if (!role) {
    Browser.msgBox("対象権限レベルが入力されていないため処理を中断します");
    return;
  }

  const sendNotificationEmail = sheets.insertPermission.getRange("権限追加_メール通知する").getValue();
  insertPermissions(email, role, sendNotificationEmail);
}
