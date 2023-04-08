/**
 * @fileoverview Helpful UI elements.
 */

/* exported uploadFile */
/* exported sheetUI */

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
  var html = HtmlService.createHtmlOutputFromFile('upload');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload File to Images folder');
}

/**
 * Hand image file upload via a base64 encoded data string.
 * @param {string} data 
 * @param {string} type 
 */
function uploadFile(data, type) {
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
    type
  );
  const fileName = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  imageBlob.setName(fileName);
  const newFile = imageFolder.createFile(imageBlob);
}
