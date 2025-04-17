/**
 * クエリを用いて特定のフォルダ配下を検索する関数
 * @param {string} folderId - 検索対象フォルダ
 * @param {string} q - 検索クエリ
 * @param {boolean} [supportsAllDrives=false] - 共有ドライブを含める
 * @param {boolean} [includeItemsFromAllDrives=false] - 共有ドライブ配下のアイテムを含める
 * @param {number} [pageSize=1000] - ページサイズ
 * @param {string} [pageToken] - ページトークン
 * @param {string} [orderBy] - 並び替え
 * @param {boolean} [useDomainAdminAccess=false] - ドメイン管理者アクセスの使用
 * @returns {Array} 検索結果のファイル一覧
 */
function searchItems(folderId, q, supportsAllDrives = false, includeItemsFromAllDrives = false, fields, pageSize = 1000, pageToken, orderBy = "folder desc", useDomainAdminAccess = false) {
  let optionalArgs = {
    fields,
    pageSize,
    pageToken,
    orderBy,
    useDomainAdminAccess,
    supportsAllDrives,
    includeItemsFromAllDrives,
  };

  const oldOptionalArgs = props.getProperty("searchOptionalArgs");
  if (oldOptionalArgs && Browser.msgBox("検索途中のクエリが見つかりました。再開しますか？", Browser.Buttons.YES_NO) === "yes") {
    optionalArgs = JSON.parse(oldOptionalArgs);
  } else if (q) {
    optionalArgs.q = q;
  } else if (folderId) {
    optionalArgs.q = `'${folderId}' in parents`;
  } else {
    optionalArgs.q = `'root' in parents`;
  }

  let fileList = [];
  do {
    props.setProperty("searchOptionalArgs", JSON.stringify(optionalArgs));
    const result = Drive.Files.list(optionalArgs);
    const { files, nextPageToken } = result;
    fileList = fileList.concat(files);
    sheets.search.getRange("検索_ステータス_クエリ").setValue(`取得件数: ${fileList.length}アイテム / クエリ: ${optionalArgs.q}`);
    optionalArgs.pageToken = nextPageToken;
  } while (optionalArgs.pageToken);
  props.setProperty("searchOptionalArgs", "");

  if (includeItemsFromAllDrives && /permissions/.test(fields)) {
    fileList = fileList.map((f) => {
      const p = getPermissions(f.id);
      return { ...f, ...p };
    });
  }

  return fileList;
}

/**
 * 指定されたファイルのIDを受け取って対象の権限を取得する関数
 * @param {string} fileId - ファイルのID
 * @param {boolean} [supportsAllDrives=true] - 共有ドライブを含める
 * @returns {Array} 権限一覧
 */
function getPermissions(fileId, supportsAllDrives = true) {
  return Drive.Permissions.list(fileId, {
    fields: "permissions(id,emailAddress,role,type)",
    supportsAllDrives,
  });
}

/**
 * 共有ドライブリストの取得
 * @returns {Array} 共有ドライブ一覧
 */
function getDrives() {
  const optionalArgs = { pageSize: 100, fields: DRIVE_FIELDS };
  let driveList = [];
  do {
    const result = Drive.Drives.list(optionalArgs);
    const { drives, nextPageToken } = result;
    driveList = driveList.concat(drives);
    optionalArgs.pageToken = nextPageToken;
  } while (optionalArgs.pageToken);

  driveList = driveList.map((d) => {
    const p = getPermissions(d.id);
    d.createdTime = d.getCreatedTime();
    d.restrictions = d.getRestrictions();
    return { ...d, ...p };
  });

  return driveList;
}

/**
 * 指定されたファイルのID、メールアドレス、権限レベルを受け取って対象の権限を削除する関数
 * @param {string} fileId - ファイルのID
 * @param {string} emailAddress - メールアドレス
 * @param {string} role - 権限レベル
 */
function removePermission(fileId, emailAddress, role) {
  try {
    const fields = "permissions(id,emailAddress,role)";
    const permissions = Drive.Permissions.list(fileId, { fields }).permissions;
    const permissionId = permissions.find((permission) => {
      return permission.emailAddress === emailAddress && (permission.role === role || role === "any");
    })?.id;

    if (permissionId) {
      Drive.Permissions.remove(fileId, permissionId);
    } else {
      return "権限なし";
    }
  } catch (e) {
    return e;
  }
  return "成功";
}

/**
 * 指定されたファイルのID、メールアドレス、権限レベルを受け取って対象の権限を追加する関数
 * @param {string} fileId - ファイルのID
 * @param {string} emailAddress - メールアドレス
 * @param {string} role - 権限レベル
 * @param {boolean} [sendNotificationEmail=false] - メールによる通知の有無
 */
function insertPermission(fileId, emailAddress, role, sendNotificationEmail = false) {
  try {
    Drive.Permissions.create({ type: "user", role, emailAddress }, fileId);
  } catch (e) {
    return e;
  }
  return "成功";
}

/**
 * mimeTypeを名前に変換
 * @param {string} type - mimeType
 * @returns {string} ファイル種別名
 */
function getFileTypeName(type) {
  if (type === "application/vnd.google-apps.folder") return "フォルダ";
  if (type === "application/vnd.google-apps.spreadsheet") return "スプレッドシート";
  if (/image\//.test(type)) return "画像";
  if (/text\//.test(type)) return "テキスト";
  if (type === "application/pdf") return "PDF";
  if (type === "application/zip") return "ZIP";
  if (type === "application/pdf") return "PDF";
  if (type === "application/x-dosexec") return "実行ファイル";
  if (type === "application/vnd.google-apps.document") return "ドキュメント";
  if (type === "application/vnd.google-apps.presentation") return "スライド";
  if (type === "application/vnd.google-apps.form") return "フォーム";
  if (type === "application/vnd.google-apps.shortcut") return "ショートカット";
  if (type === "application/vnd.google-apps.jam") return "Jamboard";
  if (/excel\//.test(type)) return "Excel";
  if (/audio\//.test(type)) return "音声";
  if (/video\//.test(type)) return "動画";
  if (type === "application/vnd.google-apps.script") return "Google Apps Script";
  if (type === "application/vnd.google-apps.site") return "Google サイト";
  if (type === "application/vnd.google-apps.photo") return "Google フォト";
  if (type === "application/vnd.google-apps.map") return "Google マップ";
  if (type === "application/json") return "JSON";
  return `その他(${type})`;
}
