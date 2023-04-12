/**
 * @fileoverview Helpful UI elements.
 */

/* exported uploadFile */
/* exported openUploadDialog */
/* exported sheetUI */

const ERROR_MISSING_VALUE = 'Selected cell was empty. Please select a cell with a string value.';

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
  const html = HtmlService.createHtmlOutputFromFile('upload');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload File to Images folder');
}

/**
 * Hand image file upload via a base64 encoded data string.
 * @param {string} data
 * @param {string} type
 */
function uploadFile(data, type) {
  const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  if (activeCell && activeCell.getValue() && activeCell.getValue().length > 0) {

    const fileName = activeCell.getValue();

    const thisFileId = SpreadsheetApp.getActive().getId();
    const thisFile = DriveApp.getFileById(thisFileId);

    const parents = thisFile.getParents();
    let f;
    while (parents.hasNext()) {
      f = parents.next();
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
  } else {
    throw new Error(ERROR_MISSING_VALUE);
  }

}
