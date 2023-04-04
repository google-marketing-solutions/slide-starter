/**
 * Google AppScript File
 * @fileoverview Script used in the UX Starter project to automate UX audits.
 *
 * UX Starter V7 - 04/04/23
 */

// Error messages
// const ERROR_NO_SPREADSHEET = 'UX Starter must be attached to a spreadsheet.'
// const ERROR_NO_DECK = 'There was a problem opening the generated deck.';
// const ERROR_CREATING_IMAGES = 'There was a problem creating the image
// mockups.'; const ERROR_MULTIPLE_FOLDERS = 'Please ensure there is only one
// folder named ';

// Warning messages
const WARNING_NO_IMAGES = 'No image found for criteria id ';
const WARNING_MULTIPLE_IMAGES = 'No image found for criteria id ';

/* exported applyCustomStyle */
/* exported onOpen */
/* exported parseFieldsAndCreateSlide */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  loadConfiguration();
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    {
      name: 'Load configuration',
      functionName: 'loadConfiguration',
    },
    {
      name: 'Filter criteria only',
      functionName: 'filterAndSortData',
    },
    {
      name: 'Generate deck',
      functionName: 'createDeckFromDatasources',
    },
  ];
  spreadsheet.addMenu('Katalyst', menuItems);
}

function createSlidesUX(deck, insightDeck, slideLayout) {
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));
  filterAndSortData();
  const values = spreadsheet.getFilter().getRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (spreadsheet.isRowHiddenByFilter(i + 1)) {
      continue;
    }
    const row = values[i];
    parseFieldsAndCreateSlide(deck, insightDeck, slideLayout, row);
  }
}


/**
 * Parses the fields contained on the incoming row from the spreadsheet into
 * some specific information fields, and then creates the slide using GAS.
 *
 * @param {string} deck Id of the generated deck that will contain the recos
 * @param {!Presentation} insightDeck Reference to the generated deck
 * @param {!Layout} recommendationSlideLayout The template layout
 * @param {!Array<string>} row Array of strings with information from the
 *     spreadsheet
 */
function parseFieldsAndCreateSlide(
    deck, insightDeck, recommendationSlideLayout, row) {
  const criteriaIdIndex =
      documentProperties.getProperty('UX_CRITERIA_ID_COLUMN') - 1;
  const criteriaNameIndex = documentProperties.getProperty('TITLE_COLUMN') - 1;
  const criteriaAppliesIndex =
      documentProperties.getProperty('SUBTITLE_COLUMN') - 1;
  const criteriaProblemStatementIndex =
      documentProperties.getProperty(
          'UX_RECOMMENDATIONS_PROBLEM_STATEMENT_ROW') -
      1;
  const criteriaSolutionStatementIndex =
      documentProperties.getProperty(
          'UX_RECOMMENDATIONS_SOLUTION_STATEMENT_ROW') -
      1;
  const criteriaImageMockupIndex =
      documentProperties.getProperty('UX_IMAGE_MOCKUP_COLUMN') - 1;
  const criteriaDefaultImageUrl =
      documentProperties.getProperty('UX_DEFAULT_IMAGE_MOCKUP');
  const criteriaInsightSlidesIndex =
      documentProperties.getProperty('INSIGHT_SLIDE_ID_COLUMN') - 1;

  const criteriaId = row[criteriaIdIndex];
  const criteria = row[criteriaNameIndex];
  const applicable =
      `Applies for: ${row[criteriaAppliesIndex].split(',').join(',')}`;
  const description = row[criteriaProblemStatementIndex] + '\n' +
      row[criteriaSolutionStatementIndex];
  const imageMockup =
      (row[criteriaImageMockupIndex] === '' ? criteriaDefaultImageUrl :
                                              row[criteriaImageMockupIndex]);
  const insights = row[criteriaInsightSlidesIndex].split(',');
  const folder = DriveApp.getFileById(SpreadsheetApp.getActive().getId())
                     .getParents()
                     .next();
  const clientImage = retrieveClientImage(folder, criteriaId);
  createRecommendationSlideGAS(
      deck, recommendationSlideLayout, criteria, applicable, description,
      imageMockup, clientImage);
  if (insights.length > 0) {
    appendInsightSlides(deck, insightDeck, insights);
  }
}

/**
 * Creates the slides programmatically using the SlidesApp from AppScript:
 * It first creates a new slide with the specified layout, it populates the
 * placeholders with the corresponding values passed as parameter.
 * In case there are insight slide ids in the insights array, it will
 * instantiate those slides and append them to the created slide
 *
 * @param {!Presentation} deck Id of the generated deck that will contain the
 *     recos
 * @param {!Layout} recommendationSlideLayout The template layout
 * @param {string} criteria The name of the criteria used as title
 * @param {string} applicable A list of pages where this criteria is applicable
 * @param {string} description The description of the failing criteria
 * @param {string} imageMockup The URL of the image corresponding to the mockup
 * @param {string} clientImage The image file of the client screenshot
 */
function createRecommendationSlideGAS(
    deck, recommendationSlideLayout, criteria, applicable, description,
    imageMockup, clientImage) {
  const slide = deck.appendSlide(recommendationSlideLayout);
  if (deck.getMasters().length > 1) {
    deck.getMasters()[deck.getMasters().length - 1].remove();
  }

  const titlePlaceholder =
      slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  const subtitlePlaceholder =
      slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
  const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);

  const titleRange = titlePlaceholder.asShape().getText();
  titleRange.setText(criteria);

  const subtitleRange = subtitlePlaceholder.asShape().getText();
  subtitleRange.setText(applicable);

  const bodyRange = bodyPlaceholder.asShape().getText();
  bodyRange.setText(description);

  const baseShape = retrieveShape(slide, 'best-practice');
  const clientShape = retrieveShape(slide, 'client-mockup');

  slide.insertImage(
      imageMockup, baseShape.getLeft(), baseShape.getTop(),
      baseShape.getWidth(), baseShape.getHeight());
  slide.insertImage(
      clientImage, clientShape.getLeft(), clientShape.getTop(),
      clientShape.getWidth(), clientShape.getHeight());
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
 * @param {string} criteriaId String corresponding to the criteria id.
 * @return {?*} Image file for the screenshot or a string url for the default
 */
function retrieveClientImage(folder, criteriaId) {
  const searchQuery = `title contains '${criteriaId}'
      and mimeType contains 'image'`;
  const files = folder.searchFiles(searchQuery);
  let file = null;

  if (files.hasNext()) {
    file = files.next();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(WARNING_NO_IMAGES + criteriaId);
  }

  if (files.hasNext()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
        WARNING_MULTIPLE_IMAGES + criteriaId);
  }

  if (file === null) {
    file = PropertiesService.getDocumentProperties().getProperty(
        'UX_DEFAULT_IMAGE_MOCKUP');
  }

  return file;
}

/**
 * Applies any extra operations to the deck based on the specifics of the audit
 *
 * @param {string} newDeckId Id of the generated deck that will contain the
 *     recos
 */
function applyCustomStyle(newDeckId) {
  const deck = SlidesApp.openById(newDeckId);
  const insightDeck =
      SlidesApp.openById(documentProperties.getProperty('INSIGHTS_DECK_ID'));
  const endSlideId = documentProperties.getProperty('END_SLIDE_ID');
  const endSlide = insightDeck.getSlideById(endSlideId.trim());
  deck.appendSlide(endSlide, SlidesApp.SlideLinkingMode.NOT_LINKED);
}

// ----- Sustainability benchmark

function parseFieldsAndCreateSlideSustainability(
    deck, insightDeck, slideLayout) {
  const presentationId = deck.getId();
  // Preparing the data and adding it into the chart
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('RECOMMENDATIONS_SHEET'));
  filterAndSortData(spreadsheet);
  const values = spreadsheet.getFilter().getRange().getValues();
  const chartSheetName = documentProperties.getProperty('DATA_SOURCE_SHEET');
  buildReadinessAnalysis(spreadsheet, values, chartSheetName);

  // Retrieving and inserting the chart
  const chartSheet = SpreadsheetApp.getActive().getSheetByName(chartSheetName);
  const spreadsheetId = SpreadsheetApp.getActive().getId();
  const sheetChartId = chartSheet.getCharts()[0].getChartId();
  const slide = deck.appendSlide(slideLayout);
  if (deck.getMasters().length > 1) {
    deck.getMasters()[deck.getMasters().length - 1].remove();
  }
  const slidePageId = slide.getObjectId();
  const slideShape = retrieveShape(slide, 'chart-location');
  deck.saveAndClose();
  replaceSlideShapeWithSheetsChart(
      presentationId, spreadsheetId, sheetChartId, slidePageId, slideShape);
}

function buildReadinessAnalysis(spreadsheet, values, chartSheetName) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const policyNamesListString =
      documentProperties.getProperty('CATEGORY_NAMES_LIST');
  const policyNamesList =
      policyNamesListString.split(',').map(item => item.trim());
  const policyColumnIndex =
      documentProperties.getProperty('POLICY_MAPPING_COLUMN') - 1;
  const policyValuesList = new Array(policyNamesList.length).fill(0);
  const policyTotalList = new Array(policyNamesList.length).fill(0);

  for (let row of values) {
    if (!row[policyColumnIndex]) {
      continue;
    }
    let rowPolicyArray =
        row[policyColumnIndex].split(',').map(item => item.trim());
    for (let policyName of rowPolicyArray) {
      let policyZeroIndex = policyNamesList.indexOf(policyName);
      policyTotalList[policyZeroIndex]++;
      let rowZeroIndex = values.indexOf(row);
      if (spreadsheet.isRowHiddenByFilter(rowZeroIndex + 1)) {
        continue;
      }
      policyValuesList[policyZeroIndex]++;
    }
  }

  const partialValuesRange = '\'' + chartSheetName + '\'!' +
      documentProperties.getProperty('PARTIAL_RESULTS_RANGE');
  const totalValuesRange = '\'' + chartSheetName + '\'!' +
      documentProperties.getProperty('TOTAL_RESULTS_RANGE');
  SpreadsheetApp.getActive().getRangeByName(partialValuesRange).setValues([
    policyValuesList
  ]);
  SpreadsheetApp.getActive().getRangeByName(totalValuesRange).setValues([
    policyTotalList
  ]);
}