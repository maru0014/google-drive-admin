/**
 * オブジェクト配列をテーブル配列に変換
 * @param {Array} obj オブジェクト配列
 * @return {Array} テーブル化された配列
 */
function object2Array(obj) {
  let table = [];
  for(const key in obj){
    table.push(obj[key]);
  }
  return table;
}

/**
 * オブジェクト配列をテーブル配列に変換
 * @param {Array} obj オブジェクト配列
 * @return {Array} テーブル化された配列
 */
function object2Table(obj) {
  const headers = [Object.keys(obj[0])];
  const body = obj.map((e) => Object.values(e));
  const table = headers.concat(body);

  return table;
}

/**
 * テーブル配列をオブジェクト配列に変換
 * @param {Array} table テーブル配列
 * @return {Array} オブジェクト化された配列
 */
function table2Object(table) {
  const arry = [];
  for (let i = 1; i < table.length; i++) {
    const obj = {};
    for (let ii = 0; ii < table[i].length; ii++) {
      obj[table[0][ii]] = table[i][ii];
    }
    arry.push(obj);
  }
  return arry;
}

/**
 * ネストされたオブジェクトをフラットにする
 * @param {Object} obj ネストされたオブジェクト
 * @return {Object} フラット化されたオブジェクト
 */
function flattenObj(obj) {
  const result = {};
  for (const key in obj) {
    const value = obj[key];
    if (value !== null && typeof value === "object") {
      const flatObj = flattenObj(value);
      for (const subKey in flatObj) {
        result[`${key}.${subKey}`] = flatObj[subKey];
      }
    } else {
      result[key] = value;
    }
  }

  return result;
}

/**
 * Google Workspaceの管理者権限をもっているかどうかを判定
 * @return {Boolean} 管理者かどうか
 */
function isAdminUser() {
  let isAdmin = false;
  try {
    const currentUser = Session.getActiveUser();
    isAdmin = AdminDirectory.Users.get(currentUser.getEmail()).isAdmin;
  } catch (e) {
    return false;
  }
  return isAdmin;
}

