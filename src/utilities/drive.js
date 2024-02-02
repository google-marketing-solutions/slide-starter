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
 * Checks whether the provided ID is a valid presentation ID.
 *
 * @param {string} deckId The ID of the presentation to check.
 * @return {boolean} True if the ID is a valid presentation ID, false
 *     otherwise.
 */
function isPresentationId(deckId) {
  try {
    const file = DriveApp.getFileById(deckId);
    if (file.getMimeType() === 'application/vnd.google-apps.presentation') {
      return true;
    } else {
      return false;
    }
  } catch (error) {
    return false;
  }
}

/**
 * Retrieves the final image value based on the provided raw value.
 * @param {string|undefined} rawValue - The raw value representing the image
 *     source.
 * @return {string|Blob} The final image URL, blob, or default image URL.
 */
function getImageBlobFromFolder(rawValue) {
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
 * Below are the exports required for the linter.
 * This is necessary because AppsScript doesn't support modules.
 */
/* exported isPresentationId */
/* exported getImageBlobFromFolder */
