// Usage: レジュームデータの管理を行うクラス
class Resume {
  /**
   * コンストラクタ
   * @param {SpreadSheet.Sheet} sheet - データ保存先シート
   */
  constructor(sheet) {
    this.sheet = sheet;

    // 実行開始時間
    this.startTime = new Date();
    // トリガー登録済みかどうか
    this.registeredResume = false;
    // 有料アカウント: 29分, 無料アカウント: 5分
    this.targetSeconds = this.isPaidUser() ? 1740 : 300;
  }

  /**
   * タスクが完了したかどうかを判定。
   * @returns {boolean} 完了したかどうか
   */
  isCompleted() {
    return searchTask.loadData().length === 0;
  }

  /**
   * 実行時間を計測します。
   * @returns {number} 実行時間（ミリ秒）
   */
  measureExecutionTime() {
    const endTime = new Date();
    const executionTime = endTime - this.startTime;
    return executionTime;
  }

  /**
   * レジュームデータの読み込み
   * @returns {Array} レジュームデータ
   */
  loadData() {
    const table = this.sheet.getDataRange().getValues();
    const data = table2Object(table);
    return data;
  }

  /**
   * レジュームデータの更新
   * @param {object} data - data
   * @returns {Array} 更新後のレジュームデータ
   */
  updateData(data) {
    const newData = object2Array(data);
    this.sheet.insertRowBefore(2);
    this.sheet.getRange(2, 1, 1, newData.length).setValues([newData]);
  }

  /**
   * レジュームデータの削除
   * @param {string} id - 削除対象のID
   * @returns {Array} 削除後のレジュームデータ
   */
  removeData(id) {
    const resumeData = this.loadData();
    for (let i = 0; i < resumeData.length; i++) {
      if (resumeData[i].id === id) {
        this.sheet.deleteRow(i + 2);
      }
    }
  }

  /**
   * レジュームデータの削除
   * @param {string} key - スクリプトプロパティキー
   */
  clearData() {
    const lastRow = this.sheet.getLastRow();
    const lastCol = this.sheet.getLastColumn();
    this.sheet.getRange(2, 1, lastRow, lastCol).clearContent();
  }

  /**
   * 自動的に処理を再開する関数
   * @param {string} functionName - 再開させる関数名
   * @param {integer} [delaySecconds=60] - トリガー登録ターゲット秒数(s)
   */
  autoResume(functionName = "main", delaySecconds = 60) {
    const elapsedSeconds = this.measureExecutionTime() / 1000;
    if (elapsedSeconds > this.targetSeconds && !this.registeredResume) {
      console.log(`autoResume TRUE: ${elapsedSeconds} > ${this.targetSeconds}`);
      this.resetTrigger(functionName, delaySecconds);
      return true;
    }
    return false;
  }

  /**
   * 自動的に処理を再開する関数
   * @param {string} functionName - 再開させる関数名
   * @param {integer} [delaySecconds60] - トリガー登録ターゲット秒数(s)
   */
  resetTrigger(functionName, delaySecconds = 60) {
    const triggers = ScriptApp.getProjectTriggers();

    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === functionName) {
        console.log(`delete trigger: ${functionName}`);
        ScriptApp.deleteTrigger(trigger);
      }
    }

    console.log(`create trigger: ${functionName}`);
    return ScriptApp.newTrigger(functionName)
      .timeBased()
      .after(1000 * delaySecconds)
      .create();
  }

  /**
   * ユーザーが有料版かどうかを判定します。
   * @returns {boolean} 有料版かどうか
   */
  isPaidUser() {
    const currentUser = Session.getActiveUser();
    const domain = currentUser.getEmail().split("@")[1];
    return domain !== "gmail.com";
  }
}
