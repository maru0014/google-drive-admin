
/**
 * 指定されたアイテムのオーナーを変更します。
 */
function getFormSettingsAll() {
  const lastRow = sheets.formSettings.getLastRow();
  if (5 > lastRow) {
    Browser.msgBox("5行目以下にデータが見つからないため処理を中断します");
    return;
  }

  for (let i = 5; i <= lastRow; i++) {
    const id = sheets.formSettings.getRange(i, 1).getValue();
    const sharingAccess = sheets.formSettings.getRange(i, 3).getValue();
    if (!id || sharingAccess) continue;

    try {
      const result = getFormSettings(id);
      const row = [[
        result.sharingAccess,
        result.isPublishingSummary,
        result.requiresLogin,  
        result.collectsEmail, 
        result.canEditResponse, 
        result.hasLimitOneResponsePerUser, 
        result.isQuiz, 
        result.editors.join(","), 
        result.editUrl]
      ];
      sheets.formSettings.getRange(i, 4, 1, 6).insertCheckboxes();
      sheets.formSettings.getRange(i, 3, 1, 9).setValues(row);
    } catch (e) {
      sheets.formSettings.getRange(i, 12).setValue(e);
    }
  }
}
