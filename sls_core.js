/* exported retrieveShape */
/* exported appendInsightSlides */
/* exported createDeckFromRecommendations */
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
 * 22/01/23
 */

// Error messages
const ERROR_MISSING_PROPERTY =
    'There\'s a missing property from the configuration.';
const ERROR_MISSING_RANGE = 'Couldn\'t find the named range in Configuration.';
const ERROR_NO_SHAPE = 'There was a problem retrieving the shape layout.';

// Properties configuration
const RANGE_NAME = 'Configuration!PROPERTIES';
const NUM_PROPERTIES = 16;

// Document properties
const documentProperties = PropertiesService.getDocumentProperties();

/**
 * Loads the configuration properties based on a named range defined on the
 * active spreadsheet and maps them to the document properties using
 * the properties service.
 *
 * @param {number} properties Number of properties to be expected
 */
function loadConfiguration(properties = NUM_PROPERTIES) {
  const range = SpreadsheetApp.getActive().getRangeByName(RANGE_NAME);
  if (!range) {
    throw new Error(ERROR_MISSING_RANGE);
  }
  const values = range.getValues();
  if (values.length < properties) {
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
 * name specified on the base template. As the API doesn't offer a direct way to
 * do this operation, it iterates over all of the existing layouts and it
 * returns the correct one once it has found a match. This function assumes that
 * the base template contains the layout name as specified on the constants.
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
  throw new Error(`There was a problem retrieving the slide layout, 
  please check the configuration tab.`);
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
 */
function filterAndSortRecommendations() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Filtering and sorting');
  const documentProperties = PropertiesService.getDocumentProperties();
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('RECOMMENDATIONS_SHEET_NAME'));

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
          documentProperties.getProperty('FAILING_CRITERIA_TEXT'));
  filter.sort(documentProperties.getProperty('SORTING_ROW'), sortingOrder)
      .setColumnFilterCriteria(
          documentProperties.getProperty('FAILED_ROW'), failingFilterCriteria);
}

/**
 * Creates a presentation based on a set of recommendations included in a
 * spreadsheet. For this, it creates a copy of a base deck, it retrieves the
 * correct template, it filters and sorts the recommendations and it creates
 * a slide operation request for each row that wasn't hidden by the filter
 * excluding the header row.
 */
function createDeckFromRecommendations() {
  loadConfiguration();
  filterAndSortRecommendations();
  const documentProperties = PropertiesService.getDocumentProperties();
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('RECOMMENDATIONS_SHEET_NAME'));
  const values = spreadsheet.getFilter().getRange().getValues();

  // TODO: Check if there are no sorted recommendations as part of that, and
  // stop the process. The problem with this is that there's no way to quickly
  // obtain the amount of results filtered

  const newDeckId = createBaseDeck();
  const recommendationSlideLayout = getTemplateLayout(newDeckId);

  const deck = SlidesApp.openById(newDeckId);
  const insightDeck =
      SlidesApp.openById(documentProperties.getProperty('INSIGHTS_DECK_ID'));

  for (let i = 1; i < values.length; i++) {
    if (spreadsheet.isRowHiddenByFilter(i + 1)) {
      continue;
    }
    const row = values[i];
    parseFieldsAndCreateSlide(
        deck, insightDeck, recommendationSlideLayout, row);
  }

  customDataInjection(newDeckId);
  applyCustomStyle(newDeckId);
}

