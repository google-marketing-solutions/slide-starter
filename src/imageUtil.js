/* exported getImagesFolder */
/* exported uploadFile */
/* exported uploadImageDialog */

/**
 * @license
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Google AppScript File
 * @fileoverview Image upload utilities.
 */

const ERROR_MISSING_VALUE = 'Please select a non-empty cell in "Image file name column". This will be used for the image file name.';
const ERROR_PARENT_FOLDER =
    'You do not have access to the parent folder of this Sheet.';
const SUCCESS_UPLOADED = 'File uploaded for: ';

/**
 * Present HTML file upload form for images.
 */
function uploadImageDialog() {
  // Get a cell with text which will be used as the file name
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeCell = sheet.getActiveCell();
  if (activeCell && activeCell.getValue() && activeCell.getValue().length > 0) {
    const criteriaName = sheet.getRange(activeCell.getRow(), 1).getValue();
    const fileName = activeCell.getValue();
    // Creates template using upload.html
    const html = HtmlService.createTemplateFromFile('upload');
    // Assigns the selected cell's text to html's criteriaName variable/scriplet
    html.criteriaName = criteriaName;
    html.fileName = fileName;
    // Executes the scriplets and converts html template to an HTMLOutput object within a modal
    SpreadsheetApp.getUi()
        .showModalDialog(html.evaluate(), 'Upload File to Images folder');
  } else {
    SpreadsheetApp.getActiveSpreadsheet()
        .toast(ERROR_MISSING_VALUE);
    throw new Error(ERROR_MISSING_VALUE);
  }
}

/**
 * Handles image file upload via a base64 encoded data string.
 *
 * @param {string} data Image file to upload.
 * @param {string} type File type.
 */
function uploadFile(data, type) {
  const imagesFolder = getImagesFolder();
  const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  const fileName = activeCell.getValue();
  const imageBlob = Utilities.newBlob(
      Utilities.base64Decode(data.split(',')[1]),
      type,
  );
  imageBlob.setName(fileName);
  imagesFolder.createFile(imageBlob);
  SpreadsheetApp.getActiveSpreadsheet()
      .toast(SUCCESS_UPLOADED + fileName);
}

/**
 * Fetches or creates "Images" folder in the parent directory.
 *
 * @return {Folder} Folder in Drive that holds images to be injected into
 * slides.
 */
function getImagesFolder() {
  const thisFileId = SpreadsheetApp.getActive().getId();
  const thisFile = DriveApp.getFileById(thisFileId);
  const parentIterator = thisFile.getParents();
  let parentFolder;
  while (parentIterator.hasNext()) {
    parentFolder = parentIterator.next();
  }
  if (!parentFolder) {
    SpreadsheetApp.getActiveSpreadsheet()
        .toast(ERROR_PARENT_FOLDER);
    throw new Error(ERROR_PARENT_FOLDER);
  }

  const imagesFolders = parentFolder.getFoldersByName('Images');
  let imagesFolder;
  if (!imagesFolders.hasNext()) {
    imagesFolder = parentFolder.createFolder('Images');
  } else {
    imagesFolder = imagesFolders.next();
  }
  return imagesFolder;
}
