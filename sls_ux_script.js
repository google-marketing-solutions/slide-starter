/**
  Copyright 2022 Google LLC

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

      https://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
 */

/**
 * Google AppScript File
 * @fileoverview Includes the core shared functions between the different
 * implementations of Slide Starter for the procedural generation of slide
 * decks.
 *
 * - loadConfiguration
 *   Populates Document Properties with the parameters retrieved from the
 *   configuration sheet on the trix
 *
 * - createBaseDeck
 *   Creates a slide deck based on a base template
 *
 * - customDataInjection
 *   Finds and replaces all placeholder strings within a slide deck
 *
 * - getTemplateLayoutId
 *   Helper function to retrieve a layout object id based on a displayName
 *   property, only accessible via advanced Slides API
 *
 * - getTemplateLayout
 *   Helper function to retrieve a layout object based on object id
 *
 * - retrieveShape
 *   Helper function to retrieve a specific shape within a slide deck based on
 *   a string
 *
 */

// Error messages
const ERROR_MISSING_PROPERTY =
    'There\'s a missing property from the configuration.';
const ERROR_MISSING_RANGE = 'Couldn\'t find the named range in Configuration.';
const ERROR_NO_SHAPE = 'There was a problem retrieving the shape layout.';

// Properties configuration
const NUM_PROPERTIES = 16;
const RANGE_NAME = 'Configuration!PROPERTIES';

/**
 * Loads the configuration properties based on a named range defined on the
 * active spreadsheet and maps them to the document properties using
 * the properties service.
 */
function loadConfiguration() {
  const range = SpreadsheetApp.getActive().getRangeByName(RANGE_NAME);
  if (!range) {
    throw new Error(ERROR_MISSING_RANGE);
  }
  const values = range.getValues();
  if (values.length < NUM_PROPERTIES) {
    throw new Error(ERROR_MISSING_PROPERTY);
  }
  for (const value of values) {
    documentProperties.setProperty(value[0], String(value[1]));
  }
}

/**
 * Copies a template deck based on the id specified on the configuration sheet.
 * It creates the deck in the same folder as the recommendations spreadsheet
 * under the assumption that this will be hosted in the vendor's drive.
 * Params are specified in document properties for ease of adjustment during
 * development.
 *
 * @return {string} Id of the copied deck
 */
function createBaseDeck() {
  const parentFolder =
      DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
          .getParents()
          .next();
  const templateDeck =
      DriveApp.getFileById(documentProperties.getProperty('TEMPLATE_DECK_ID'));
  return templateDeck
      .makeCopy(documentProperties.getProperty('DECK_NAME'), parentFolder)
      .getId();
}

/**
 * Finds and replaces all placeholder strings within a slide deck. It stops
 * processing whenever it finds the first empty row within the sheet.
 * @param {string} newDeckId Id of the new slide deck that has
 *     been generated
 */
function customDataInjection(newDeckId) {
  const presentation = SlidesApp.openById(newDeckId);

  SpreadsheetApp.getActiveSpreadsheet().toast('Autofilling strings');
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DICTIONARY_SHEET_NAME'));

  const dictionary = sheet.getDataRange().offset(1, 0).getValues();

  for (const row of dictionary) {
    if (!row[0]) break;
    presentation.replaceAllText(row[0], row[1]);
  }
}

/**
 * Retrieves the template layout id from the presentation based on the template
 * name specified on the base template. As the API doesn't offer a direct way to do
 * this operation, it iterates over all of the existing layouts and it returns
 * the correct one once it has found a match. This function assumes that the
 * base template contains the layout name as specified on the constants.
 *
 * @param {string} presentationId Id of the new slide deck that has
 *     been generated
 * @return {string} Id of the layout matched by defined name
 */
function getTemplateLayoutId(presentationId) {
  const layouts = Slides.Presentations.get(presentationId).layouts;
  for (const layout of layouts) {
    if (layout.layoutProperties.displayName ===
        documentProperties.getProperty('LAYOUT_NAME')) {
      return layout.objectId;
    }
  }
  throw new Error(
      'There was a problem retrieving the slide layout, please check the configuration tab.');
}

/**
 * Retrieves a Layout object based on a template layout name defined on the
 * properties sheet in the presentation.
 *
 * @param {string} presentationId Id of the new slide deck that has
 *     been generated
 * @return {?Layout} Layout object matched by defined name if found or null
 */
function getTemplateLayout(presentationId) {
  const layouts = SlidesApp.openById(presentationId).getLayouts();
  const layoutId = getTemplateLayoutId(presentationId);
  for (const layout of layouts) {
    if (layout.getObjectId() === layoutId) {
      return layout;
    }
  }
  return null;
}


/**
 * Returns a Shape object from a slide layout based on a string match. Given API
 * limitations, it must iterate over all existing shapes in the slide to
 * retrieve the desired one, then match based on the find function on its
 * enclosed TextRange.
 *
 * @param {!Slide} slide The slide to find the Shape.
 * @param {string} typeString The string to match in order to find the shape
 * @return {!Shape} Shape found, or empty string
 */
function retrieveShape(slide, typeString) {
  for (const shape of slide.getLayout().getShapes()) {
    const shapeText = shape.getText();
    if (shapeText.find(typeString).length) {
      return shape;
    }
  }
  throw new Error(ERROR_NO_SHAPE);
}