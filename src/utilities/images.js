/**
 * @license
 * Copyright 2024 Google LLC
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
 * @fileoverview Helpful UI elements.
 */

/* exported uploadFile */
/* exported openUploadDialog */
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
    const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
    if (activeCell && activeCell.getValue() && activeCell.getValue().length > 0) {
      const criteriaName =
          SpreadsheetApp.getActiveSheet().getRange(activeCell.getRow(), 1);
      const html = HtmlService.createTemplateFromFile('image_upload');
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
  