/**
 * レジュームデータをもとに検索を再開する
 * @param {string} resumeData - 親フォルダID
 * @param {string} q - クエリ
 * @param {boolean} allDrives - 共有ドライブも対象とする
 * @param {boolean} getPermissions - 権限情報も取得する
 * @param {boolean} getSubfolders - サブフォルダー配下のアイテムも再帰的に取得する
 */
function resumeSearch(resumeData, q, allDrives, getPermissions, getSubfolders) {
  if (resumeData?.hasOwnProperty("triggerUid")) {
    // トリガーから呼び出された場合は自動的に読み込む
    q = sheets.search.getRange("検索_クエリ").getValue();
    allDrives = sheets.search.getRange("検索_共有ドライブ").getValue();
    getSubfolders = q ? false : sheets.search.getRange("検索_サブフォルダも検索").getValue();
    getPermissions = sheets.search.getRange("検索_権限出力").getValue();
    resumeData = searchTask.loadData();
  }
  console.log("resumeSearch", JSON.stringify(resumeData));

  for (const data of resumeData) {
    const { id, path } = data;
    if (id) {
      search(id, q, allDrives, allDrives, getPermissions, getSubfolders, path);
    }
  }

  if (searchTask.isCompleted()) {
    sheets.search.getRange("検索_ステータス").setValue(`Done`);
    sheets.search.sort(1, true);
  }
}

/**
 * 親フォルダのIDを受け取って子アイテムをすべて取得する
 * @param {string} parentId - 親フォルダID
 * @param {string} q - クエリ
 * @param {boolean} allDrives - 共有ドライブも対象とする
 * @param {boolean} getPermissions - 権限情報も取得する
 * @param {boolean} getSubfolders - サブフォルダー配下のアイテムも再帰的に取得する
 * @param {string} path - フォルダパス
 */
function search(parentId, q, allDrives = false, getPermissions, getSubfolders, path = "") {
  if (searchTask.registeredResume) {
    // レジュームトリガーセット済みの場合は処理を中断する
    sheets.search.getRange("検索_ステータス").setValue(`AutoResumeを待機中 ... 60秒後に処理を再開予定です`);
    return;
  }
  searchTask.autoResume("resumeSearch");

  // フォルダ情報を取得
  const parent = parentId ? Drive.Files.get(parentId, {
    supportsAllDrives: allDrives,
    includeItemsFromAllDrives: allDrives,
  }) : "";
  const currentPath = path || parent?.name || "";

  // アイテムを検索
  sheets.search.getRange("検索_ステータス").setValue(`Search ... ${currentPath}`);
  const fields = getPermissions ? FILE_AND_PERMISSION_FIELDS : FILE_FIELDS;
  const items = searchItems(parentId, q, allDrives, allDrives, fields);

  // 直下のアイテムを出力
  const itemDataTable = parseItemsData(items, getPermissions, parent, currentPath);
  console.log("get items:", currentPath, itemDataTable.length);
  if (itemDataTable.length) {
    const lastRow = sheets.search.getLastRow();
    sheets.search.getRange(lastRow + 1, 1, itemDataTable.length, itemDataTable[0].length).setValues(itemDataTable);
  }
  searchTask.removeData(parentId);

  // サブフォルダも取得する
  if (getSubfolders) {
    // フォルダのみ抽出
    const folders = items.filter((item) => item.mimeType === "application/vnd.google-apps.folder");

    // フォルダをレジュームデータに登録
    for (const folder of folders) {
      const folderPath = `${currentPath}/${folder.name}`;
      searchTask.updateData({ id: folder.id, path: folderPath });
    }

    // 再帰的にコール
    for (const folder of folders) {
      const folderPath = `${currentPath}/${folder.name}`;
      console.log("get subitems:", folderPath);
      search(folder.id, q, allDrives, getPermissions, getSubfolders, folderPath);
    }
  }

  sheets.search.getRange("検索_ステータス_クエリ").setValue(``);
}

/**
 * 検索結果をスプレッドシート出力用の二次元配列に変換
 * @param {array} items - 検索結果item配列
 * @param {boolean} getPermissions - 権限情報も取得する
 * @param {object} parent - 親フォルダ
 * @param {string} path - フォルダパス
 * @returns {Array} 二次元配列
 */
function parseItemsData(items, getPermissions, parent, path) {
  const table = [];

  for (const item of items) {
    const { name, id, mimeType, permissions, shortcutDetails } = item;
    const mimeTypeName = getFileTypeName(mimeType);
    const targetMimeTypeName = shortcutDetails ? getFileTypeName(shortcutDetails?.targetMimeType) : "";
    const targetId = shortcutDetails?.targetId;

    if (getPermissions) {
      for (const permission of permissions || []) {
        const { type, emailAddress, role } = permission;
        const permissionTarget = ["user", "group"].includes(type) ? emailAddress : type;
        table.push([path, id, name, mimeTypeName, permissionTarget, role, parent?.id, parent?.name, targetId, targetMimeTypeName]);
      }
    } else {
      table.push([path, id, name, mimeTypeName, null, null, parent?.id, parent?.name, targetId, targetMimeTypeName]);
    }
  }

  return table;
}


/**
 * 共有ドライブリストを出力
 */
function searchDrives() {
  const drives = getDrives();
  const table = [];

  for (const drive of drives) {
    const { name, id, createdTime, restrictions, permissions} = drive;
    for (const permission of permissions || []) {
      const { emailAddress, role } = permission;
      const { driveMembersOnly, domainUsersOnly, copyRequiresWriterPermission } = restrictions;
      table.push([id, name, createdTime, emailAddress, role, driveMembersOnly, domainUsersOnly, copyRequiresWriterPermission]);
    }
  }

  if (table.length) {
    const lastRow = sheets.drives.getLastRow();
    sheets.drives.getRange(lastRow + 1, 1, table.length, table[0].length).setValues(table);
  }
}
