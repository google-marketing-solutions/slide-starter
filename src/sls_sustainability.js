/* exported createSlidesUX */
/* exported parseFieldsAndCreateSlideSustainability */

/**
 * @license
 * Copyright 2023 Google LLC
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *  https://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Google AppScript File
 * @fileoverview Script used in the UX Starter project to automate UX audits.
 *
 * UX Starter V7 - 04/04/23
 */

/* exported onEdit */
/* exported customStyledTextFields */
/* exported applyCustomStyle */

/**
 * Object whose keys represent Core Web Vital metrics and values are Arrays that
 * contain the low & high thresholds for that metric. Used in coloring the table
 * for CrUX CWV data.
 */
const CWV = {
  LCP: [2500, 4000],
  FID: [100, 300],
  INP: [100, 300],
  CLS: [0.1, 0.25],
};

const cwvTextType = {
  'CRUX_FID': CWV.FID,
  'CRUX_CLS': CWV.CLS,
  'CRUX_INP': CWV.INP,
  'CRUX_LCP': CWV.LCP,
};

/**
 * Object whose keys are colors and values are arrays of RGB values in decimal.
 * Used in coloring the table for CrUX CWV data.
 */
const COLORS = {
  GREEN: '#34A853', // Good
  YELLOW: '#FBBC04', // Needs Improvement
  RED: '#EA4335', // Poor
  WHITE: '#F8F9FA', // None
};

/**
 * A special function that runs whenever a change on the spreadsheet is detected
 * In this case we will use it to handle the modification of adding or removing
 * types of audit into Katalyst.
 *
 * As a temporary solution while we manage access to an API to record telemetry
 * we will look at changes at the telemetry sheet to make sure that an
 * installable trigger can be executed as a user with permissions to store the
 * telemetry data from the different sheets on a file that other users would not
 * have access to.
 *
 * @param {!Event} e The onEdit event.
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
  const targetSheetName = e.range.offset(0, 1).getValue();
  const sheet = e.source.getSheetByName(targetSheetName);
  (value) ? sheet.showSheet() : sheet.hideSheet();
}

// ----- Performance pre-collection function

// ----- Performance post-collection function

// ----- Performance post-slide function

/**
 * This is a post slide-creation function hook intended to apply custom styles
 * to extra text fields that were created. It receives any extra arguments that
 * were defined as an object in the configuration sheet (or source, in the
 * future). Since only the standard placeholders retain style after creation
 * (other text fields are copied based on transform of the shape), this is
 * necessary to ensure custom styling.
 *
 * @param {!Slide} slide Slide to be modified
 * @param {!Array<string>} row Row of information from data source
 * @param {!Array<string>} postSlideFunctionArgs Extra information passed down
 *     through the configuration sheet
 */
function customStyledTextFields(slide, row, postSlideFunctionArgs) {
  const textFields = postSlideFunctionArgs;
  const textShapesArray = textFields.shapes;
  const textColumnsArray = textFields.columns;

  for (let i = 0; i < textShapesArray.length; i++) {
    const shapeId = textShapesArray[i];
    const column = textColumnsArray[i];
    if (shapeId && column) {
      const textShape = retrieveShape(slide, shapeId);
      const textValue = row[column - 1].toString();
      if (textValue) {
        const newTextBox = slide.insertTextBox(
            textValue, textShape.getLeft(), textShape.getTop(),
            textShape.getWidth(), textShape.getHeight());

        // Centers text
        newTextBox.getText().getParagraphStyle().setParagraphAlignment(
            SlidesApp.ParagraphAlignment.CENTER);

        // Sets style and CWV-speficic formatting
        const textStyle = newTextBox.getText().getTextStyle();
        const shapeStyle = textShape.getText().getTextStyle();
        textStyle.setFontFamily(shapeStyle.getFontFamily());
        textStyle.setFontSize(shapeStyle.getFontSize());
        // textStyle.setFontWeight(shapeStyle.getFontWeight());

        if (Object.prototype.hasOwnProperty.call(cwvTextType, shapeId)) {
          const textColor = colorForCWV(cwvTextType[shapeId], textValue);
          textStyle.setForegroundColor(textColor);

          const formattedCWV =
              textFormatForCWV(cwvTextType[shapeId], textValue);
          newTextBox.getText().setText(formattedCWV);
        } else {
          textStyle.setForegroundColor('#1f2023');
        }
      }
    }
  }
}

/**
 * Determines a color based on if a value is a Good, Needs Improvement or Poor
 * range for a given metric.
 *
 * @param {!Array} range Array with a low and high threshold for a CWV metric
 * @param {number} value Number indicating the metric score
 * @return {!Array} Array of RBG values in decimal form
 */
function colorForCWV([lowThreshold, highThreshold], value) {
  if (!value) {
    return COLORS.WHITE;
  } else if (value <= lowThreshold) {
    return COLORS.GREEN;
  } else if (value < highThreshold) {
    return COLORS.YELLOW;
  } else {
    return COLORS.RED;
  }
}

/**
 * Converts a Core Web Vitals value to a human-readable string.
 *
 * @param {!CWV} cwv The Core Web Vitals metric.
 * @param {number} value The value of the metric.
 * @return {string} The human-readable string.
 */
function textFormatForCWV(cwv, value) {
  switch (cwv) {
    case CWV.LCP:
      return value / 1000 + 's';
    case CWV.FID:
      return value + 'ms';
    case CWV.CLS:
      return value;
    case CWV.INP:
      return value + 'ms';
    default:
      return value;
  }
}

// ----- Post deck creation styling function
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
 * @param {!GoogleAppsScript.Slides.Presentation} deck The slide deck to which
 *     the slide should be appended.
 * @param {!GoogleAppsScript.Slides.Presentation} insightDeck The slide deck
 *     that contains the insights slides to be used as context.
 * @param {!GoogleAppsScript.Slides.PageElement} slideLayout The layout of the
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
 * @param {!GoogleAppsScript.Spreadsheet.Sheet} spreadsheet The sheet containing
 *     the recommendations data.
 * @param {!Array} values The array of values returned by the filter on the
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
