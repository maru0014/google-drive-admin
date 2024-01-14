/**
 * 指定された権限を削除します。
 * @param {string} email - 対象ユーザーメールアドレス
 * @param {string} role - 対象権限レベル
 */
function removePermissions(email, role) {
  const lastRow = sheets.removePermission.getLastRow();
  if (5 > lastRow) {
    Browser.msgBox("6行目以下にデータが見つからないため処理を中断します");
    return;
  }

  for (let i = 6; i <= lastRow; i++) {
    const itemId = sheets.removePermission.getRange(i, 1).getValue();
    const status = sheets.removePermission.getRange(i, 3).getValue();
    if (itemId && status === "") {
      const result = removePermission(itemId, email, role);
      sheets.removePermission.getRange(i, 3).setValue(result);
    }
  }
}

/**
 * 指定された権限を作成します。
 * @param {string} email - 対象ユーザーメールアドレス
 * @param {string} role - 対象権限レベル
 * @param {boolean} [sendNotificationEmail=false] - メールによる通知の有無
 */
function insertPermissions(email, role, sendNotificationEmail = false) {
  const lastRow = sheets.insertPermission.getLastRow();
  if (5 > lastRow) {
    Browser.msgBox("6行目以下にデータが見つからないため処理を中断します");
    return;
  }

  for (let i = 6; i <= lastRow; i++) {
    const itemId = sheets.insertPermission.getRange(i, 1).getValue();
    const status = sheets.insertPermission.getRange(i, 3).getValue();
    if (itemId && status === "") {
      const result = insertPermission(itemId, email, role, sendNotificationEmail);
      sheets.insertPermission.getRange(i, 3).setValue(result);
    }
  }
}
