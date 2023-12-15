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
 * @fileoverview Functions related to simplifying interfacing with Sheets or running
 * operations not supported directly by either the Apps Script or REST API.
 * This includes getting layouts by name, inserting charts, retrieving shapes, creating decks...
 */

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
      .makeCopy(
          documentProperties.getProperty('OUTPUT_DECK_NAME'), parentFolder)
      .getId();
}

/**
 * Retrieves a Layout object based on a template layout name defined on the
 * properties sheet in the presentation.
 *
 * @param {string} presentationId Id of the new slide deck that has
 *     been generated
 * @param {string} [layoutName] The name of the template layout to retrieve.
 * @return {?Layout} Layout object matched by defined name if found or null
 */
function getTemplateLayout(presentationId, layoutName = null) {
  const layouts = SlidesApp.openById(presentationId).getLayouts();
  const layoutId = getTemplateLayoutId(presentationId, layoutName);
  for (const layout of layouts) {
    if (layout.getObjectId() === layoutId) {
      return layout;
    }
  }
  return null;
}

/**
 * Retrieves the template layout id from the presentation based on the template
 * name specified on the base template. As the API doesn't offer a direct way to
 * do this operation, it iterates over all of the existing layouts and it
 * returns the correct one once it has found a match. This function assumes that
 * the base template contains the layout name as specified on the constants.
 *
 * @param {string} presentationId The ID of the new slide deck that has been
 *  generated.
 * @param {string=} layoutName (optional) The name of the layout to match.
 * @return {string} The ID of the layout matched by the defined name.
 * @throws {Error} If there is a problem retrieving the slide layout.
 */
function getTemplateLayoutId(presentationId, layoutName = null) {
  const layouts = Slides.Presentations.get(presentationId).layouts;
  const nameToMatch =
      layoutName ? layoutName : documentProperties.getProperty('LAYOUT_NAME');
  for (const layout of layouts) {
    if (layout.layoutProperties.displayName === nameToMatch) {
      return layout.objectId;
    }
  }
  throw new Error(`There was a problem retrieving the slide layout, 
  please check the configuration tab.`);
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
  throw new Error(ERROR_NO_SHAPE + ' ' + typeString);
}

/**
 * Adds text to the specified placeholder on the slide.
 *
 * @param {!Slide} slide The slide to add the text to.
 * @param {!SlidesApp.PlaceholderType} placeholderType The type of placeholder
 *     to add the text to.
 * @param {string} text The text to add to the placeholder.
 * @param {string} defaultValue The default text to add to the placeholder if
 *     `text` is empty.
 */
function addTextToPlaceholder(slide, placeholderType, text, defaultValue) {
  const placeholder = slide.getPlaceholder(placeholderType).asShape().getText();
  if (text && text.length > 0) {
    placeholder.setText(text);
  } else {
    placeholder.setText(defaultValue);
  }
}

/**
 * Finds and replaces all placeholder strings within a slide deck. It stops
 * processing whenever it finds the first empty row within the sheet.
 * @param {string} newDeckId Id of the new slide deck that has
 *     been generated
 */
//TODO: Refactor name - Something more descriptive "DeckWideTextReplacement"
function customDataInjection(newDeckId) {
  const presentation = SlidesApp.openById(newDeckId);

  SpreadsheetApp.getActiveSpreadsheet().toast('Autofilling strings');
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DICTIONARY_SHEET_NAME'),
  );

  const dictionary = sheet.getDataRange().offset(1, 0).getValues();

  for (const row of dictionary) {
    if (!row[0]) break;
    presentation.replaceAllText(row[0], row[1]);
  }
}

/**
 * Creates a subsection header slide within a specified deck
 * and appends it.
 *
 * @param {string} deckId Object identifier for the slide deck
 * @param {!Layout} layout Layout object relative to the header slide
 * @param {string} title Name of the section of the audit
 */
function createSlideWithTitle(deckId, layout, title) {
  const deck = SlidesApp.openById(deckId);
  const slide = deck.appendSlide(layout);
  const titlePlaceholder = slide.getPlaceholder(
      SlidesApp.PlaceholderType.TITLE,
  );
  const titleRange = titlePlaceholder.asShape().getText();
  titleRange.setText(title);
}

/**
 * Embed a Sheets chart (indicated by the spreadsheetId and sheetChartId) onto
 * a page in the presentation. Setting the linking mode as 'LINKED' allows the
 * chart to be refreshed if the Sheets version is updated.
 * We don't use the objectId when creating the Sheets chart, but the API
 * requires it, so we use the value of the current full datetime to ensure there
 * are no duplicates.
 *
 * @param {string} presentationId
 * @param {string} spreadsheetId
 * @param {string} sheetChartId
 * @param {string} slidePageId
 * @param {string} slideChartShape
 * @return {*}
 */
function replaceSlideShapeWithSheetsChart(
  presentationId, spreadsheetId, sheetChartId, slidePageId, slideChartShape) {
  const chartHeight = slideChartShape.getInherentHeight();
  const chartWidth = slideChartShape.getInherentWidth();
  const chartTransform = slideChartShape.getTransform();
  const requiredForAPIButUnused = new Date().toDateString();
  const requests = [
    {
      createSheetsChart: {
        objectId: requiredForAPIButUnused,
        spreadsheetId: spreadsheetId,
        chartId: sheetChartId,
        linkingMode: 'LINKED',
        elementProperties: {
          pageObjectId: slidePageId,
          size: {
            width: {magnitude: chartHeight, unit: 'PT'},
            height: {magnitude: chartWidth, unit: 'PT'},
          },
          transform: {
            scaleX: chartTransform.getScaleX(),
            scaleY: chartTransform.getScaleY(),
            translateX: chartTransform.getTranslateX(),
            translateY: chartTransform.getTranslateY(),
            unit: 'PT',
          },
        },
      },
    },
  ];

  // Execute the request.
  try {
    const batchUpdateResponse =
        Slides.Presentations.batchUpdate({requests: requests}, presentationId);
    console.log('Added a linked Sheets chart with ID: %s', presentationId);
    slideChartShape.remove();
    return batchUpdateResponse;
  } catch (err) {
    console.log('Failed with error: %s', err);
  }
}