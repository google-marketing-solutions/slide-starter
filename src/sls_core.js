/* exported retrieveShape */
/* exported appendInsightSlides */
/* exported createDeckFromDatasources */
/* exported replaceSlideShapeWithSheetsChart*/
/**
 * Google AppScript File
 * @fileoverview Includes the core shared functions between the different
 * implementations of Slide Starter for the procedural generation of slide
 * decks.
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
 * 14/04/23
 */

// Error messages
const ERROR_MISSING_RANGE = 'Couldn\'t find the named range in Configuration.';
const ERROR_NO_SHAPE = 'There was a problem retrieving the shape layout.';

// Properties configuration
const RANGE_NAME = 'Configuration!PROPERTIES';

// Document properties
const documentProperties = PropertiesService.getDocumentProperties();

/**
 * Loads the configuration properties based on a named range defined on the
 * active spreadsheet and maps them to the document properties using
 * the properties service.
 *
 * @param {string} rangeName Optional name of the range to use
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
  const titlePlaceholder =
      slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
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
 *
 * @param {Sheet} sheet - The sheet to apply the filter and sort to. Defaults to
 *     the active sheet.
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
  const requests = [{
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
  }];

  // Execute the request.
  try {
    const batchUpdateResponse =
        Slides.Presentations.batchUpdate({requests: requests}, presentationId);
    console.log('Added a linked Sheets chart with ID: %s', presentationChartId);
    slideChartShape.remove();
    return batchUpdateResponse;
  } catch (err) {
    // TODO (Developer) - Handle exception
    console.log('Failed with error: %s', err.error);
    console.log('Failed with error: %s', err);
  }
}

/**
 * Gets a function by name.
 *
 * @param {string} functionName The name of the function to get.
 * @return {Function} The function with the given name.
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
  return url.startsWith('http://') || url.startsWith('https://');
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
 * multiple images, the last one is selected.
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

// --- Katalyst loops

/**
 * Creates a new Slides deck based on the data sources specified in the document
 * properties. Uses the specified base deck as a template, and applies custom
 * styling to the new deck.
 */
function createDeckFromDatasources() {
  const documentProperties = PropertiesService.getDocumentProperties();
  loadConfiguration();

  const newDeckId = createBaseDeck();

  const datasourceString = documentProperties.getProperty('DATA_SOURCE_SHEET');
  const datasourcesArray =
      datasourceString.split(',').map((item) => item.trim());

  const sectionLayoutName =
      documentProperties.getProperty('SECTION_LAYOUT_NAME');
  let sectionLayout;
  if (sectionLayoutName && sectionLayoutName.length > 0) {
    sectionLayout = getTemplateLayout(newDeckId, sectionLayoutName);
  }

  for (const datasource of datasourcesArray) {
    if (sectionLayout) {
      createHeaderSlide(deckId, sectionLayout, datasource);
    }
    prepareDependenciesAndCreateSlides(datasource, newDeckId);
  }

  const dictionarySheetName =
      documentProperties.getProperty('DICTIONARY_SHEET_NAME');
  if (dictionarySheetName && dictionarySheetName.length > 0) {
    customDataInjection(newDeckId);
  }

  applyCustomStyle(newDeckId);
}

/**
 * Prepares the dependencies and creates slides for the given datasource.
 *
 * @param {string} datasource The name of the datasource.
 * @param {string} newDeckId The ID of the new deck to create slides in.
 */
function prepareDependenciesAndCreateSlides(datasource, newDeckId) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const datasourceConfiguration =
      '\'Configuration_' + datasource + '\'!PROPERTIES';
  loadConfiguration(datasourceConfiguration);

  const deck = SlidesApp.openById(newDeckId);
  const recommendationSlideLayout = getTemplateLayout(newDeckId);

  let insightDeck;
  const insightsDeckId = documentProperties.getProperty('INSIGHTS_DECK_ID');
  if (insightsDeckId && insightsDeckId.length > 0) {
    insightDeck =
        SlidesApp.openById(documentProperties.getProperty('INSIGHTS_DECK_ID'));
  }

  createSlidesForDatasource(deck, insightDeck, recommendationSlideLayout);
}

/**
 * If a custom function is specified in the configuration tab, the custom
 * function is called instead. Otherwise t creates either a single slide or a
 * collection slide based on the check in config.
 *
 * @param {Presentation} deck - The Slides deck where the new slide(s) will be
 *     created.
 * @param {Presentation} insightDeck - Extra deck to pull insight slides,
 *     retrieved only once.
 * @param {Layout} slideLayout - The slide layout to use for the new slide(s).
 *
 */
function createSlidesForDatasource(deck, insightDeck, slideLayout) {
  const customFunctionName = documentProperties.getProperty('CUSTOM_FUNCTION');
  if (customFunctionName && customFunctionName.length > 0) {
    getFunctionByName(customFunctionName)(deck, insightDeck, slideLayout);
  } else {
    const isSingleSlide =
        documentProperties.getProperty('SINGLE_VALUE') == 'TRUE';
    if (isSingleSlide) {
      createSingleSlide(deck, insightDeck, slideLayout);
    } else {
      createCollectionSlide(deck, insightDeck, slideLayout);
    }
  }
}

/**
 * Creates a collection slide based on data from a Google Sheets data source
 * using the specified deck, insight deck, and slide layout. Filters and sorts
 * the data, and creates a slide for each row that passes the filter criteria.
 *
 * @param {Presentation} deck - The Slides deck where the new slide(s) will be
 *     created.
 * @param {Presentation} insightDeck - The Slides deck where the insight slide
 *     will be created (if applicable).
 * @param {Layout} slideLayout - The slide layout to use for the new slide(s).
 *
 */
function createCollectionSlide(deck, insightDeck, slideLayout) {
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));
  filterAndSortData();
  const values = spreadsheet.getFilter().getRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (spreadsheet.isRowHiddenByFilter(i + 1)) {
      continue;
    }
    const row = values[i];
    parseFieldsAndCreateCollectionSlide(deck, insightDeck, slideLayout, row);
  }
}

/**
 * Creates a single slide in a Google Slides presentation.
 * The function first gets the spreadsheet that contains the data for the slide.
 * Then, it creates a new slide in the presentation using the specified layout.
 * If there is a master slide, the function removes it. Next, the function
 * fetches the title, subtitle, and body text for the slide from the
 * spreadsheet. It then sets the title, subtitle, and body text for the slide.
 * Finally, the function fetches the image shapes and ranges for the slide from
 * the spreadsheet. If there are image shapes and ranges, the function adds the
 * images to the slide.
 *
 *
 * @param {SlidesApp.Presentation} deck The presentation to add the slide to.
 * @param {SlidesApp.InsightDeck} insightDeck The insight deck that contains the
 *     data for the slide.
 * @param {SlidesApp.SlideLayout} slideLayout The layout to use for the slide.
 *
 * @return {void}
 */
function createSingleSlide(deck, insightDeck, slideLayout) {
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));

  const slide = deck.appendSlide(slideLayout);
  if (deck.getMasters().length > 1) {
    deck.getMasters()[deck.getMasters().length - 1].remove();
  }

  // Fetch fields
  const titleRange = documentProperties.getProperty('TITLE_RANGE');
  if (titleRange && titleRange.length > 0) {
    const title = spreadsheet.getRange(titleRange).getValue();
    const slideTitlePlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    const slideTitle = slideTitlePlaceholder.asShape().getText();
    slideTitle.setText(title);
  }


  const subtitleRange = documentProperties.getProperty('SUBTITLE_RANGE');
  if (subtitleRange && subtitleRange.length > 0) {
    const subtitle = spreadsheet.getRange(subtitleRange).getValue();
    const slideSubtitlePlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    const slideSubtitle = slideSubtitlePlaceholder.asShape().getText();
    slideSubtitle.setText(subtitle);
  }

  const bodyRange = documentProperties.getProperty('BODY_RANGE');
  if (bodyRange && bodyRange.length > 0) {
    const body = spreadsheet.getRange(bodyRange).getValue();
    const slideBodyPlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    const slideBody = slideBodyPlaceholder.asShape().getText();
    slideBody.setText(body);
  }

  const imageShapesArray = documentProperties.getProperty('IMAGE_SHAPES')
      .split(',')
      .map((item) => item.trim());
  const imageRangesArray = documentProperties.getProperty('IMAGE_RANGES')
      .split(',')
      .map((item) => item.trim());

  if (imageShapesArray && imageShapesArray.length > 0) {
    for (let i = 0; i < imageShapesArray.length; i++) {
      const shapeId = imageShapesArray[i];
      const range = imageRangesArray[i];

      if (shapeId && range) {
        const imageShape = retrieveShape(slide, shapeId);
        const imageValue = spreadsheet.getRange(range).getValue();
        slide.insertImage(
            imageValue, imageShape.getLeft(), imageShape.getTop(),
            imageShape.getWidth(), imageShape.getHeight());
      }
    }
  }
}


/**
 * Creates a collection slide based on a slide layout and data from a specified
 * row in a Google Sheet.
 * @param {GoogleAppsScript.Slides.Presentation} deck - The slide deck to add
 *     the slide to.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} insightDeck - The sheet
 *     containing the data to populate the slide.
 * @param {GoogleAppsScript.Slides.Layout} slideLayout - The layout to use for
 *     the slide.
 * @param {!Array<string>} row Array of strings with information from the
 *     spreadsheet
 */
function parseFieldsAndCreateCollectionSlide(
    deck, insightDeck, slideLayout, row) {
  const slide = deck.appendSlide(slideLayout);
  if (deck.getMasters().length > 1) {
    deck.getMasters()[deck.getMasters().length - 1].remove();
  }

  // Add title
  const titleColumn = documentProperties.getProperty('TITLE_COLUMN');
  if (titleColumn && titleColumn.length > 0) {
    const title = row[titleColumn - 1];
    const slideTitlePlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    const slideTitle = slideTitlePlaceholder.asShape().getText();
    slideTitle.setText(title);
  }

  // Add subtitle
  const subtitleColumn = documentProperties.getProperty('SUBTITLE_COLUMN');
  if (subtitleColumn && subtitleColumn.length > 0) {
    const subtitle = row[subtitleColumn - 1];
    const slideSubtitlePlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    const slideSubtitle = slideSubtitlePlaceholder.asShape().getText();
    slideSubtitle.setText(subtitle);
  }

  // Add body
  const bodyColumn = documentProperties.getProperty('BODY_COLUMN');
  if (bodyColumn && bodyColumn.length > 0) {
    const body = row[bodyColumn - 1];
    const slideBodyPlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    const slideBody = slideBodyPlaceholder.asShape().getText();
    slideBody.setText(body);
  }

  // Add images
  const imageShapesArray = documentProperties.getProperty('IMAGE_SHAPES')
      .split(',')
      .map((item) => item.trim());
  const imageColumnsArray = documentProperties.getProperty('IMAGE_COLUMNS')
      .split(',')
      .map((item) => item.trim());

  if (imageShapesArray && imageShapesArray.length > 0) {
    for (let i = 0; i < imageShapesArray.length; i++) {
      const shapeId = imageShapesArray[i];
      const column = imageColumnsArray[i];

      if (shapeId && column) {
        const imageShape = retrieveShape(slide, shapeId);
        let imageValue = row[column - 1];
        if (!imageValue) {
          imageValue = documentProperties.getProperty('DEFAULT_IMAGE_URL');
        } else if (!isValidImageUrl(imageValue)) {
          const folder =
              DriveApp.getFileById(SpreadsheetApp.getActive().getId())
                  .getParents()
                  .next();
          const imageName = imageValue;
          imageValue = retrieveImageFromFolder(folder, imageName);
        }
        slide.insertImage(
            imageValue, imageShape.getLeft(), imageShape.getTop(),
            imageShape.getWidth(), imageShape.getHeight());
      }
    }
  }

  // Add insight slides
  const insightSlidesColumn =
      documentProperties.getProperty('INSIGHT_SLIDE_ID_COLUMN');
  const insights =
      row[insightSlidesColumn - 1].split(',').map((item) => item.trim());
  if (insights.length > 0) {
    appendInsightSlides(deck, insightDeck, insights);
  }
}
