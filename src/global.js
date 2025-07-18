// load sheets
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheets = {
  drives: ss.getSheetByName("ドライブ"),
  search: ss.getSheetByName("アイテム検索"),
  searchFolders: ss.getSheetByName("search_folders"),
  searchQuery: ss.getSheetByName("search_query"),
  insertPermission: ss.getSheetByName("権限追加"),
  removePermission: ss.getSheetByName("権限削除"),
  changeOwner: ss.getSheetByName("オーナー変更"),
  drives: ss.getSheetByName("共有ドライブ一覧"),
  formSettings: ss.getSheetByName("フォーム設定確認"),
}

// load script properties
const props = PropertiesService.getScriptProperties();

// resume class
const searchTask = new Resume(sheets.searchFolders);

// constants
const FILE_AND_PERMISSION_FIELDS = "nextPageToken, files(id,shortcutDetails,name,mimeType,permissions,parents,driveId)";
const FILE_FIELDS = "nextPageToken, files(id,shortcutDetails,name,mimeType,parents,driveId)";
const DRIVE_FIELDS = "nextPageToken, drives(id,name,createdTime,restrictions)";
