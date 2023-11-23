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

/* exported replaceSlideShapeWithSheetsChart */
/* exported retrieveShape */
/* exported getTemplateLayout */
/* exported createBaseDeck */
/* exported loadConfiguration */
/* exported addTextToPlaceholder */
/* exported retrieveImageFromFolder */
/* exported getFunctionByName */
/* exported isValidImageUrl */
/* exported createHeaderSlide */
/* exported customDataInjection */
/* exported shouldCreateCollectionSlide */
/* exported appendInsightSlides */
/* exported filterAndSortData */
/* exported colorForCWV */
/* exported getImageValue */

/*
Global redefined here to prevent access errors from the tests.
This will be addressed with
*/
documentProperties = PropertiesService.getDocumentProperties();





/**
 * Loads the configuration properties based on a named range defined on the
 * active spreadsheet and maps them to the document properties using
 * the properties service.
 *
 * @param {string=} rangeName Optional name of the range to use
 */
function loadConfiguration(rangeName = RANGE_NAME) {
  const range = SpreadsheetApp.getActive().getRangeByName(rangeName);
  if (!range) {
    throw new Error(ERROR_MISSING_RANGE);
  }
  const values = range.getValues();
  for (const value of values) {
    documentProperties.setProperty(value[0], String(value[1]));
  }
}



// Drive
/**
 * Returns a file, which is assumed to be an image file, for a criteria client
 * screenshot which should be named after a criteria id. Any formats are
 * considered for the query, but it is assumed that the file will be an image.
 *
 * This file is retrieved from a folder created programmatically which is
 * assumed to exist.
 *
 * If no such file has been found, the function returns the url of the default
 * image mockup, which behaves analogously to an image file.
 *
 * There are warnings sent out (currently as a toast on the spreadsheet) in this
 * case, and also in case that multiple image files are found. When finding
 * multiple images, the first one is selected.
 *
 * @param {!Folder} folder The folder where images are being stored.
 * @param {string} imageName Name of the image to be found.
 * @return {?*} Image file for the screenshot or a string url for the default
 */
function retrieveImageFromFolder(folder, imageName) {
  const searchQuery = `title contains '${imageName}'
  and mimeType contains 'image'`;
  const files = folder.searchFiles(searchQuery);
  let file = null;

  if (files.hasNext()) {
    file = files.next();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(WARNING_NO_IMAGES + imageName);
  }

  if (files.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
        WARNING_MULTIPLE_IMAGES + imageName);
  }

  if (file === null) {
    file = PropertiesService.getDocumentProperties().getProperty(
        'DEFAULT_IMAGE_URL');
  }

  return file;
}

/**
 * Checks if the provided image value is a base64 encoded image.
 * @param {string} imageValue - The string value representing the image.
 * @return {boolean} True if the image is base64 encoded, False otherwise.
 */
function isBase64Image(imageValue) {
  return imageValue.match(/^data:image\/([a-z]+);base64,/i);
}

/**
 * Decodes a base64 encoded image string and returns a blob.
 * @param {string} imageValue - The base64 encoded image string.
 * @return {Blob} A blob object containing the decoded image data.
 */
function decodeBase64Image(imageValue) {
  const match = imageValue.match(/^data:image\/([a-z]+);base64,/i);
  const imageType = match[1];
  const imageBase64 = imageValue.split(',')[1];
  const decodedImage = Utilities.base64Decode(imageBase64);
  return Utilities.newBlob(decodedImage, MimeType[imageType.toUpperCase()]);
}

/**
 * Retrieves the final image value based on the provided raw value.
 * @param {string|undefined} rawValue - The raw value representing the image
 *     source.
 * @return {string|Blob} The final image URL, blob, or default image URL.
 */
function getImageValue(rawValue) {
  let imageValue =
      rawValue || documentProperties.getProperty('DEFAULT_IMAGE_URL');
  if (isBase64Image(imageValue)) {
    imageValue = decodeBase64Image(imageValue);
  } else if (!isValidImageUrl(imageValue)) {
    const folder = DriveApp.getFileById(SpreadsheetApp.getActive().getId())
        .getParents()
        .next();
    imageValue = retrieveImageFromFolder(folder, imageValue);
  }
  return imageValue;
}

// Generic

/**
 * Gets a function by name.
 *
 * @param {string} functionName The name of the function to get.
 * @return {!Function} The function with the given name.
 * @throws {Error} If the function name is not alphanumeric.
 */
function getFunctionByName(functionName) {
  const alphanumericRegex = /^[a-zA-Z0-9]+$/;
  if (!alphanumericRegex.test(functionName)) {
    throw new Error('Function name not alphanumeric');
  }
  return new Function(`return ${functionName};`)();
}

/**
 * Checks if the given URL is a valid image URL.
 *
 * @param {string} url The URL to check.
 * @return {boolean} Whether the URL is a valid image URL.
 */
function isValidImageUrl(url) {
  // TODO: Fix issue #39 Check if it's an image
  return url.startsWith('http://') || url.startsWith('https://');
}

// Katalyst helpers


/**
 * Determines whether the script should create a slide for this row in the
 * collection based on whether title, subtitle, or body column have been
 * defined. Only one of them should be in order to create a slide.
 *
 * @return {boolean} Whether the script should create a slide for that row in
 *     the collection
 */
function shouldCreateCollectionSlide() {
  const titleColumn = documentProperties.getProperty('TITLE_COLUMN');
  const subtitleColumn = documentProperties.getProperty('SUBTITLE_COLUMN');
  const bodyColumn = documentProperties.getProperty('BODY_COLUMN');
  return (
    (titleColumn && titleColumn.length > 0) ||
      (subtitleColumn && subtitleColumn.length > 0) ||
      (bodyColumn && bodyColumn.length > 0));
}

/**
 * Appends insight slides by reference to the generated deck
 *
 * @param {!Presentation} deck Reference to the generated deck
 * @param {!Presentation} insightDeck Reference to the generated deck
 * @param {!Array<string>} insights Array of slide ids for extended insights
 */
function appendInsightSlides(deck, insightDeck, insights) {
  for (const insightSlideId of insights) {
    if (insightSlideId === '') {
      continue;
    }
    const insightSlide = insightDeck.getSlideById(insightSlideId.trim());
    if (insightSlide === null) {
      continue;
    }
    deck.appendSlide(insightSlide, SlidesApp.SlideLinkingMode.NOT_LINKED);
    if (deck.getMasters().length > 1) {
      deck.getMasters()[deck.getMasters().length - 1].remove();
    }
  }
}

/**
 * Retrieves the active spreadsheet,removes the current filter if it exists,
 * applies new filter based on criteria, and sorts by a specified column in the
 * trix.
 *
 * @param {!Sheet=} sheet - The sheet to apply the filter and sort to. Defaults
 *     to the active sheet.
 */
function filterAndSortData(sheet = undefined) {
  SpreadsheetApp.getActiveSpreadsheet().toast('Filtering and sorting');
  const documentProperties = PropertiesService.getDocumentProperties();
  if (!sheet) {
    sheet = SpreadsheetApp.getActive().getSheetByName(
        documentProperties.getProperty('DATA_SOURCE_SHEET'));
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const previousFilter = sheet.getRange(1, 1, lastRow, lastColumn).getFilter();

  if (previousFilter !== null) {
    previousFilter.remove();
  }

  let sortingOrder = false;
  if (documentProperties.getProperty('SORTING_ORDER')) {
    sortingOrder = Boolean(documentProperties.getProperty('SORTING_ORDER'));
  }

  const filter = sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
  const filterColumn = documentProperties.getProperty('FILTER_COLUMN');
  if (filterColumn && filterColumn.length > 0) {
    const failingFilterCriteria =
        SpreadsheetApp.newFilterCriteria().whenTextContains(
            documentProperties.getProperty('FILTER_TEXT_VALUE'));
    filter.sort(documentProperties.getProperty('SORTING_COLUMN'), sortingOrder)
        .setColumnFilterCriteria(
            documentProperties.getProperty('FILTER_COLUMN'),
            failingFilterCriteria);
  }
}

// ---- Other helpers

/**
 * Determines a color based on if a value is a Good, Needs Improvement or Poor
 * range for a given metric.
 *
 * @param {!Array} range Array with a low and high threshold for a CWV metric
 * @param {number} value Number indicating the metric score
 * @return {!Array} Array of RBG values in decimal form
 */
function colorForCWV([lowThreshold, highThreshold], value) {
  if (!value.trim()) {
    return COLORS.WHITE;
  } else if (value <= lowThreshold) {
    return COLORS.GREEN;
  } else if (value < highThreshold) {
    return COLORS.YELLOW;
  } else {
    return COLORS.RED;
  }
}
