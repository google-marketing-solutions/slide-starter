/* exported createSlidesUX */
/* exported parseFieldsAndCreateSlideSustainability */
/**
 * Google AppScript File
 * @fileoverview Script used in the UX Starter project to automate UX audits.
 *
 * UX Starter V7 - 04/04/23
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  loadConfiguration();
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    {
      name: 'Generate deck',
      functionName: 'createDeckFromDatasources',
    },
  ];
  spreadsheet.addMenu('Katalyst', menuItems);
}

/**
 * A special function that runs whenever a change on the spreadsheet is detected
 * In this case we will use it to handle the modification of adding or removing 
 * types of audit into Katalyst.
 */
function onEdit(e) {
  const currentSheetName = e.source.getActiveSheet().getName();
  if (currentSheetName != 'Audit Start') {
    return;
  }
  const currentColumn = e.range.columnStart;
  if (currentColumn != 1) { 
    return;
  }
  const value = e.range.getValue();
  const targetSheetName = e.range.offset(0,1).getValue();
  const sheet = e.source.getSheetByName(targetSheetName);
  const _ = (value) ? sheet.showSheet() : sheet.hideSheet();
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
      SlidesApp.openById(documentProperties.getProperty('END_SLIDE_DECK_ID'));
  const endSlideId = documentProperties.getProperty('END_SLIDE_ID');
  const endSlide = insightDeck.getSlideById(endSlideId.trim());
  deck.appendSlide(endSlide, SlidesApp.SlideLinkingMode.NOT_LINKED);
}

// ----- Sustainability benchmark

/**
 * Parses the fields and creates a sustainability slide. This function prepares
 * the data from the recommendations sheet and builds a chart that is inserted
 * into the slide. The chart is taken from a sheet specified in the constants.
 * The slide is appended to the presentation specified by the deck parameter.
 *
 * @param {GoogleAppsScript.Slides.Presentation} deck The slide deck to which
 *     the slide should be appended.
 * @param {GoogleAppsScript.Slides.Presentation} insightDeck The slide deck that
 *     contains the insights slides to be used as context.
 * @param {GoogleAppsScript.Slides.PageElement} slideLayout The layout of the
 *     slide to be created.
 */
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

/**
 * Builds the readiness analysis chart based on the recommendations data. This
 * function is called by parseFieldsAndCreateSlideSustainability and is
 * responsible for building the chart that is inserted into the slide. It reads
 * the configuration data from the document properties and uses the
 * recommendations sheet to calculate the values to be displayed in the chart.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} spreadsheet The sheet containing
 *     the recommendations data.
 * @param {Array} values The array of values returned by the filter on the
 *     sheet.
 * @param {string} chartSheetName The name of the sheet that contains the chart
 *     to be used.
 */
function buildReadinessAnalysis(spreadsheet, values, chartSheetName) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const policyNamesListString =
      documentProperties.getProperty('CATEGORY_NAMES_LIST');
  const policyNamesList =
      policyNamesListString.split(',').map((item) => item.trim());
  const policyColumnIndex =
      documentProperties.getProperty('POLICY_MAPPING_COLUMN') - 1;
  const policyValuesList = new Array(policyNamesList.length).fill(0);
  const policyTotalList = new Array(policyNamesList.length).fill(0);

  for (const row of values) {
    if (!row[policyColumnIndex]) {
      continue;
    }
    const rowPolicyArray =
        row[policyColumnIndex].split(',').map((item) => item.trim());
    for (const policyName of rowPolicyArray) {
      const policyZeroIndex = policyNamesList.indexOf(policyName);
      policyTotalList[policyZeroIndex]++;
      const rowZeroIndex = values.indexOf(row);
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
    policyValuesList,
  ]);
  SpreadsheetApp.getActive().getRangeByName(totalValuesRange).setValues([
    policyTotalList,
  ]);
}
