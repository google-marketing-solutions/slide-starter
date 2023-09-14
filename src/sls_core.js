/* exported createDeckFromDatasource */
/* exported loadConfiguration */
/* exported retrieveShape */

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
 * @fileoverview Includes the core shared functions between the different
 * implementations of Slide Starter for the procedural generation of slide
 * decks.
 * - createDeckFromDatasource
 *   Creates an audit deck based on configuration settings.
 *
 * - loadConfiguration
 *   Populates Document Properties with the parameters retrieved from the
 *   configuration sheet on the trix.
 *
 * - filterAndSortData
 *   Uses catalog sheet to filter and sort recommendations to include in final
 *   deck.
 *
 * - createBaseDeck
 *   Creates a slide deck based on a base template.
 *
 * - customDataInjection
 *   Finds and replaces all placeholder strings within a slide deck.
 *
 * - replaceText
 *   Finds and replaces strings within a slide deck.
 *
 * - addAppendixDeck
 *   Adds an entire deck of slides to the end of the final deck.
 *
 * - retrieveShape
 *   Fetches a shape from a slide or slide layout.
 */

// Error messages
const ERROR_MISSING_RANGE = 'Couldn\'t find the named range in Configuration.';
const ERROR_NO_SHAPE = 'There was a problem retrieving the shape layout.';

// Named range of configuration properties
const CONFIG_PROPERTIES = 'Configuration!PROPERTIES';

const documentProperties = PropertiesService.getDocumentProperties();

/**
 * Creates a presentation based on a set of recommendations included in a
 * spreadsheet.
 */
function createDeckFromDatasource() {
  loadConfiguration();
  filterAndSortData();
  const newDeckId = createBaseDeck();
  const documentProperties = PropertiesService.getDocumentProperties();
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));
  const recommendationsRange = spreadsheet.getFilter().getRange();
  const recommendationsValues = recommendationsRange.getValues();

  // Skip the header row after filtering recommendations sheet
  for (let i = 1; i < recommendationsValues.length; i++) {
    // Passing criteria will be hidden by the filter
    if (spreadsheet.isRowHiddenByFilter(i + 1)) {
      continue;
    }
    const row = recommendationsValues[i];
    // i+1 because range includes header
    appendInsightDeck(newDeckId, row, i+1, recommendationsRange);
  }
  customDataInjection(newDeckId);
  addAppendixDeck(newDeckId);
  applyCustomStyle(newDeckId);
}

/**
* Loads the configuration properties based on a named range defined on the
* active spreadsheet and maps them to the document properties using
* the properties service.
*
* @param {string} rangeName Optional name of the range to use.
*/
function loadConfiguration(rangeName = CONFIG_PROPERTIES) {
  const range = SpreadsheetApp.getActive().getRangeByName(rangeName);
  if (!range) {
    throw new Error(ERROR_MISSING_RANGE);
  }
  const values = range.getValues();
  for (const row of values) {
    documentProperties.setProperty(row[0], String(row[1]));
  }
}

/**
 * Retrieves the active spreadsheet,removes the current filter if it exists,
 * applies new filter based on criteria, and sorts by a specified column in the
 * trix.
 */
function filterAndSortData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Filtering and sorting');
  const documentProperties = PropertiesService.getDocumentProperties();
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));

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
  const failingFilterCriteria =
      SpreadsheetApp.newFilterCriteria().whenTextContains(
          documentProperties.getProperty('FILTER_TEXT_VALUE'));
  filter.setColumnFilterCriteria(
      documentProperties.getProperty('FILTER_COLUMN'),
      failingFilterCriteria)
      .sort(documentProperties.getProperty('SORTING_COLUMN'), sortingOrder);
}

/**
 * Copies a template deck based on the id specified on the configuration sheet.
 * It creates the deck in the same folder as the recommendations spreadsheet
 * under the assumption that this will be hosted in the vendor's drive.
 * Params are specified in document properties for ease of adjustment during
 * development.
 *
 * @return {string} Id of the copied deck to use as the final Presentation.
 */
function createBaseDeck() {
  const parentFolder =
      DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
          .getParents()
          .next();
  const templateDeck =
      DriveApp.getFileById(documentProperties.getProperty('TEMPLATE_DECK_ID'));
  return templateDeck
      .makeCopy(
          documentProperties.getProperty('OUTPUT_DECK_NAME'), parentFolder)
      .getId();
}

/**
 * Finds and replaces all placeholder strings within a slide deck. It stops
 * processing whenever it finds the first empty row within the sheet.
 *
 * @param {!string} newDeckId Id of the new slide deck that has
 * been generated.
 */
function customDataInjection(newDeckId) {
  const presentation = SlidesApp.openById(newDeckId);
  SpreadsheetApp.getActiveSpreadsheet().toast('Autofilling strings');
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DICTIONARY_SHEET_NAME'));
  const customStrings = sheet.getDataRange().offset(1, 0).getValues();
  replaceText(presentation, customStrings);
}

/**
 * Replaces all instances of text matching find text with replace text within
 * the nestedArray.
 *
 * @param {Presentation} presentation Presentation with text to replace.
 * @param {array} nestedArray Array where each item is an array with
 *      two words: [stringToFind, replacement].
 */
function replaceText(presentation, nestedArray) {
  for (const row of nestedArray) {
    if (!row[0]) break;
    presentation.replaceAllText(row[0], row[1]);
  }
}

/**
 * Adds an appendix Presentation to the end of the Presentation with newDeckId.
 *
 * @param {!string} newDeckId Id of the deck to add the appendix to.
 */
function addAppendixDeck(newDeckId) {
  const appendixDeckId = documentProperties.getProperty('APPENDIX_DECK_ID');
  if (appendixDeckId) {
    const appendixSlides = SlidesApp.openById(appendixDeckId).getSlides();
    const thisDeck = SlidesApp.openById(newDeckId);
    for (const slide of appendixSlides) {
      thisDeck.appendSlide(slide, SlidesApp.SlideLinkingMode.NOT_LINKED);
    }
  }
}

/**
 * Gets a shape from either a slide or the slide's layout.
 *
 * @param {!Slide} slide Slide to get shape from.
 * @param {string} searchString String that the desired shape contains.
 * @param {boolean} isLayout Whether to search the layout or the slide.
 * @return {Shape} Found shape in the Slide that contains the searchText.
 */
function retrieveShape(slide, searchString, isLayout = true) {
  // Either search the slides layout shapes or the slide itself
  const layoutOrSlide = isLayout ? slide.getLayout() : slide;
  for (const shape of layoutOrSlide.getShapes()) {
    const shapeText = shape.getText();
    if (shapeText.find(searchString).length) {
      return shape;
    }
  }
  throw new Error(ERROR_NO_SHAPE);
}
