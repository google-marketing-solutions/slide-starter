/* exported appendInsightDeck */
/* exported applyCustomStyle */
/* exported onOpen */

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
 * @fileoverview UX specific functions used to automate audit generation.
 * - onOpen
 *   Special function that runs when spreadsheet is open, used to load settings
 *   and create the Katalyst menu.
 *
 * - appendInsightDeck
 *   Finds and attaches a deck of insights for a given criteria.
 *
 * - retrieveImage
 *   Finds an image from the "Images" parent folder.
 *
 * - applyCustomStyle
 *   Unused function for UX but kept in order to not throw an error in core.gs
 */


// Warning messages
const WARNING_NO_IMAGES = 'No image found for criteria id ';
const WARNING_MULTIPLE_IMAGES = 'No image found for criteria id ';


/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  loadConfiguration();
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    {
      name: 'Upload Image',
      functionName: 'uploadImageDialog',
    },
    {
      name: 'Generate deck',
      functionName: 'createDeckFromDatasource',
    },
  ];
  spreadsheet.addMenu('Katalyst', menuItems);
}


/**
 * Opens Presentation with the newDeckId and appends a deck of insight slides
 * based on the first item of the row. Will add client & best practice images.
 *
 * @param {!String} newDeckId Id of the new slide deck that has
 * been generated.
 * @param {Array} row An array of text from a spreadsheet row containing
 * [insightDeckName, unused, clientImageName, bestPracticeImageName] where
 * insightDeckName is the title of the criteria & insight deck filename, unused
 * is necessary in the sheet but not used in this function, and clientImageName
 * & bestPracticeImageName are filenames to replace the client-mockup and
 * best-practice shapes in the insight slides.
 * @param {Number} index Integer representing the row number of the
 * recommendation in the catalog sheet.
 * @param {Range} rangeRef Reference to the range of recommendations.
 */
function appendInsightDeck(newDeckId, row, index, rangeRef) {
  // Title of the criteria, column A, is the name of the deck to search for
  // File name for the client image, column B, is clientImageName
  // eslint-disable-next-line no-unused-vars
  const [insightDeckName, unused, clientImageName, bestPracticeImageName] = row;
  const currentDeck = SlidesApp.openById(newDeckId);
  const insightFolderID = documentProperties.getProperty('INSIGHTS_FOLDER_ID');
  const insightFolder = DriveApp.getFolderById(insightFolderID);
  const fileIterator = insightFolder.getFilesByName(insightDeckName);
  // There should only be one file in each folder for now
  while (fileIterator.hasNext()) {
    const insightDeckId = fileIterator.next().getId();
    const insightDeck = SlidesApp.openById(insightDeckId).getSlides();
    for (const slide of insightDeck) {
      const newSlide = currentDeck.appendSlide(
          slide,
          SlidesApp.SlideLinkingMode.NOT_LINKED,
      );
      const imagesFolder = getImagesFolder();
      try {
        // Cell in catalog sheet with the client-mockup file name in column C
        const cell = rangeRef.getCell(index, 3);
        const clientImageSrc = retrieveImage(
            imagesFolder, clientImageName, cell);
        const clientShape = retrieveShape(newSlide, 'client-mockup', false);
        const insertedClientImage = newSlide.insertImage(
            clientImageSrc, clientShape.getLeft(), clientShape.getTop(),
            clientShape.getWidth(), clientShape.getHeight());
        insertedClientImage.sendToBack();
        for (let i = 0; i < 3; i++) {
          insertedClientImage.bringForward();
        }
      } catch (e) {
        console.log(
            'Potential error uploading client image:',
            insightDeckName,
            ', error:',
            e,
        );
      }
      try {
        // Cell in catalog sheet with the client-mockup file name in column D
        const cell = rangeRef.getCell(index, 4);
        const bestPracticeImageSrc = retrieveImage(
            imagesFolder,
            bestPracticeImageName,
            cell,
        );
        const bestPracticeShape = retrieveShape(
            newSlide,
            'best-practice',
            false,
        );
        const insertedBestPracticeImage = newSlide.insertImage(
            bestPracticeImageSrc,
            bestPracticeShape.getLeft(), bestPracticeShape.getTop(),
            bestPracticeShape.getWidth(), bestPracticeShape.getHeight());
        insertedBestPracticeImage.sendToBack();
        for (let i = 0; i < 4; i++) {
          insertedBestPracticeImage.bringForward();
        }
      } catch (e) {
        console.log(
            'Potential error uploading best practice image:',
            insightDeckName,
            ', error:',
            e,
        );
      }
    }
  }
}


/**
 * Returns a file, which is assumed to be an image file, for a criteria client
 * screenshot which should be named after a criteria. Any formats are
 * considered for the query, but it is assumed that the file will be an image.
 *
 * This file is retrieved from a folder created programmatically which is
 * assumed to exist.
 *
 * If no such file has been found, the function returns the url of the default
 * image mockup, which behaves analogously to an image file.
 * There are warnings sent out (currently as a toast on the spreadsheet) in this
 * case, and also in case that multiple image files are found. When finding
 * multiple images, the last one is selected.
 *
 * @param {!Folder} folder The folder where images are being stored.
 * @param {String} criteriaId String corresponding to the image name.
 * @param {Cell} cell Reference to the cell with the name of the image.
 * @return {?*} Image file for the screenshot or a string url for the default
 */
function retrieveImage(folder, criteriaId, cell) {
  const searchQuery = `title contains '${criteriaId}'
      and mimeType contains 'image'`;
  const files = folder.searchFiles(searchQuery);
  let file = null;

  if (files.hasNext()) {
    file = files.next();
    cell.setBackground('white');
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(WARNING_NO_IMAGES + criteriaId);
    console.log(`Using default image for ${criteriaId}`);
    cell.setBackground('red');
    file = documentProperties.getProperty('UX_DEFAULT_IMAGE_MOCKUP');
  }

  if (files.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
        WARNING_MULTIPLE_IMAGES + criteriaId);
  }

  return file;
}

/**
 * Applies any extra operations to the deck based on the specifics of the audit.
 *
 * @param {!String} newDeckId Id of the new slide deck that has
 * been generated.
 */
function applyCustomStyle(newDeckId) {
  return;
}

