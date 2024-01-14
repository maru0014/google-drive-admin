// load sheets
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheets = {
  drives: ss.getSheetByName("ドライブ"),
  search: ss.getSheetByName("アイテム検索"),
  insertPermission: ss.getSheetByName("権限追加"),
  removePermission: ss.getSheetByName("権限削除"),
  changeOwner: ss.getSheetByName("オーナー変更")
}

// load script properties
const props = PropertiesService.getScriptProperties();

// resume class
const searchTask = new Resume(props, "SEARCH_RESUME_DATA");

// constants
const FILE_AND_PERMISSION_FIELDS = "nextPageToken, files(id,shortcutDetails,name,mimeType,permissions)";
const FILE_FIELDS = "nextPageToken, files(id,shortcutDetails,name,mimeType)";
