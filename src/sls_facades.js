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
/* global documentProperties */

documentProperties = PropertiesService.getDocumentProperties();

/**
 * Embed a Sheets chart (indicated by the spreadsheetId and sheetChartId) onto
 *   a page in the presentation. Setting the linking mode as 'LINKED' allows the
 *   chart to be refreshed if the Sheets version is updated.
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
  const presentationChartId = 'chart-test';
  const requests = [
    {
      createSheetsChart: {
        objectId: presentationChartId,
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
    console.log('Added a linked Sheets chart with ID: %s', presentationChartId);
    slideChartShape.remove();
    return batchUpdateResponse;
  } catch (err) {
    console.log('Failed with error: %s', err);
  }
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
  // TODO: Check if it's an image
  return url.startsWith('http://') || url.startsWith('https://');
}

// Katalyst helpers
/**
 * Creates a subsection header slide within a specified deck
 * and appends it.
 *
 * @param {string} deckId Object identifier for the slide deck
 * @param {!Layout} layout Layout object relative to the header slide
 * @param {string} title Name of the section of the audit
 */
function createHeaderSlide(deckId, layout, title) {
  const deck = SlidesApp.openById(deckId);
  const slide = deck.appendSlide(layout);
  const titlePlaceholder = slide.getPlaceholder(
      SlidesApp.PlaceholderType.TITLE,
  );
  const titleRange = titlePlaceholder.asShape().getText();
  titleRange.setText(title);
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
      documentProperties.getProperty('DICTIONARY_SHEET_NAME'),
  );

  const dictionary = sheet.getDataRange().offset(1, 0).getValues();

  for (const row of dictionary) {
    if (!row[0]) break;
    presentation.replaceAllText(row[0], row[1]);
  }
}

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
