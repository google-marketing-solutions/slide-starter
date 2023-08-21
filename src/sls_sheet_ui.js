/**
 * @fileoverview Helpful UI elements.
 */

/* exported uploadFile */
/* exported openUploadDialog */
/* exported sheetUI */

const ERROR_MISSING_VALUE = 'Please select a non-empty cell.';
const ERROR_PARENT_FOLDER =
    'You do not have access to the parent folder of this Sheet.';
const SUCCESS_UPLOADED = 'File uploaded for: ';

/**
 * Add menu items for helpful UI.
 */
function sheetUI() {
  SpreadsheetApp.getUi()
      .createMenu('Wizard')
      .addItem('Upload', 'openUploadDialog')
      .addToUi();
}

/**
 * Present HTML file Upload form.
 */
function openUploadDialog() {
  const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  if (activeCell && activeCell.getValue() && activeCell.getValue().length > 0) {
    const criteriaName =
        SpreadsheetApp.getActiveSheet().getRange(activeCell.getRow(), 1);
    const html = HtmlService.createTemplateFromFile('upload');
    html.criteriaName = criteriaName.getValue();
    SpreadsheetApp.getUi()
        .showModalDialog(html.evaluate(), 'Upload File to Images folder');
  } else {
    SpreadsheetApp.getActiveSpreadsheet()
        .toast(ERROR_MISSING_VALUE);
    throw new Error(ERROR_MISSING_VALUE);
  }
}

/**
 * Hand image file upload via a base64 encoded data string.
 * @param {string} data
 * @param {string} type
 */
function uploadFile(data, type) {
  const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  const fileName = activeCell.getValue();

  const thisFileId = SpreadsheetApp.getActive().getId();
  const thisFile = DriveApp.getFileById(thisFileId);

  const parents = thisFile.getParents();
  let f;
  while (parents.hasNext()) {
    f = parents.next();
  }

  if (!f) {
    SpreadsheetApp.getActiveSpreadsheet()
        .toast(ERROR_PARENT_FOLDER);
    throw new Error(ERROR_PARENT_FOLDER);
  }
  const imagesFolders = f.getFoldersByName('Images');
  let imageFolder;
  if (!imagesFolders.hasNext()) {
    imageFolder = f.createFolder('Images');
  } else {
    imageFolder = imagesFolders.next();
  }
  const imageBlob = Utilities.newBlob(
      Utilities.base64Decode(data.split(',')[1]),
      type,
  );
  imageBlob.setName(fileName);
  imageFolder.createFile(imageBlob);
  SpreadsheetApp.getActiveSpreadsheet()
      .toast(SUCCESS_UPLOADED + fileName);
}
