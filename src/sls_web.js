/* exported appendInsightDeck */
/* exported applyCustomStyle */
/* exported createPerfSlides */
/* exported onOpen */

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
 * @fileoverview Performance specific functions used in audit generation.
 * - onOpen
 *   Special function that runs when spreadsheet is open, used to load settings
 *   and create the Katalyst menu.
 *
 * - createPerfSlides
 *   Main function that fetches PageSpeed Insights metrics & creates audit deck.
 *
 * - appendInsightDeck
 *   Finds and attaches a deck of insights for a given criteria.
 *
 * - retrieveImage
 *   Finds an image from the "Images" parent folder.
 *
 * - applyCustomStyle
 *   Updates dynamic CWV title slide & applies formatting updates.
 *
 * - buildBackgroundCellColorTableStyleSlidesRequest
 *   Collects SlidesAPI requests for formatting updates.
 *
 * - colorForCWV
 *   Gets a color based on a CWV score.
 *
 * - colorCWVTable
 *   Generates SlidesAPI requests to color the CWV table based on metrics.
 *
 * - setHyperlinksInCWVTable
 *   Converts URL plain text to hyperlins in the CWV table.
 *
 * - truncateUrl
 *   Prevents long URLs running over one line.
 *
 * - replaceCWVSlideTitle
 *   Determines and sets the dynamic title of the CWV slide giving an overview
 *   of how all the URLs performed.
 *
 * - getUrlInsights
 *   Creates insights used to determine the dynamic title of the CWV slide.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  try {
    loadConfiguration();
    const spreadsheet = SpreadsheetApp.getActive();
    const menuItems = [
      {
        name: 'Generate audit deck',
        functionName: 'createPerfSlides',
      },
    ];
    spreadsheet.addMenu('Katalyst', menuItems);
  } catch (error) {
    throw new Error('onOpen failed: ' + error.toString());
  }
}

/**
 * Loads configuration, fetches PSI metrics for URLs, and creates slide deck.
 */
function createPerfSlides() {
  loadConfiguration();
  cloneSitesSheet();
  runBatchFromQueue();
  createDeckFromDatasource();
}

/**
 * Opens Presentation with the newDeckId and appends a deck of insight slides
 * based on the first item of the row. Will add client & best practice images.
 *
 * @param {!string} newDeckId Id of the new slide deck that has
 * been generated.
 * @param {array} row An array of text from a spreadsheet row containing
 * [insightDeckName, pages] where insightDeckName is the title of the criteria &
 * insight deck filename, and pages are the failing audited pages which will be
 * injected into slides.
 */
function appendInsightDeck(newDeckId, row) {
  const [insightDeckName, pages] = row;
  const currentDeck = SlidesApp.openById(newDeckId);
  const insightFolderID = documentProperties.getProperty('INSIGHTS_FOLDER_ID');
  const insightFolder = DriveApp.getFolderById(insightFolderID);
  const fileIterator = insightFolder.getFilesByName(insightDeckName);
  // There should only be one file in each folder for now
  while (fileIterator.hasNext()) {
    const insightDeckId = fileIterator.next().getId();
    const insightDeck = SlidesApp.openById(insightDeckId).getSlides();
    for (const slide of insightDeck) {
      currentDeck.appendSlide(slide, SlidesApp.SlideLinkingMode.NOT_LINKED);
    }
  }
  const stringsToReplace = [['{{pages}}', pages]];
  replaceText(currentDeck, stringsToReplace);
}

/**
 * Object whose keys represent Core Web Vital metrics and values are Arrays that
 * contain the low & high thresholds for that metric. Used in coloring the table
 * for CrUX CWV data.
 */
const CWV = {
  LCP: [2.5, 4],
  FID: [100, 300],
  CLS: [0.1, 0.25],
  INP: [200, 500],
};

/**
 * Object whose keys are colors and values are arrays of RGB values in decimal.
 * Used in coloring the table for CrUX CWV data.
 */
const COLORS = {
  GREEN: [.04, .80, .41], // Good
  YELLOW: [1, 0.64, 0], // Needs Improvement
  RED: [1, 0.30, 0.25], // Poor
  WHITE: [1, 1, 1], // None
};

/**
 * Applies any extra operations to the deck based on the specifics of the audit.
 *
 * @param {!string} newDeckId Id of the new slide deck that has
 * been generated.
 */
function applyCustomStyle(newDeckId) {
  replaceCWVSlideTitle(newDeckId);
  documentProperties.setProperty('SLIDES_REQUESTS', JSON.stringify([]));
  colorCWVTable(newDeckId);
  const resource = {
    requests: JSON.parse(documentProperties.getProperty('SLIDES_REQUESTS')),
  };
  Slides.Presentations.batchUpdate(resource, newDeckId);
}

/**
 * Builds a SlidesAPI request to handle the table formatting properties that
 * are not accessible via the SlidesAPI service, such as column width.
 * These requests are retrieved and stored from the document properties after
 * being flattened as JSON.
 *
 * @param {string} tableId String that identifies the table to modify.
 * @param {Number} rowIndex Number that identifies the row to modify.
 * @param {Number} columnIndex Number that identifies the column to modify.
 * @param {Array} color Array of RGB values for the table cell.
 */
function buildBackgroundCellColorTableStyleSlidesRequest(
    tableId, rowIndex, columnIndex, color) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const requests =
      JSON.parse(documentProperties.getProperty('SLIDES_REQUESTS'));

  requests.push({
    updateTableCellProperties: {
      objectId: tableId,
      tableRange: {
        location: {
          rowIndex: rowIndex,
          columnIndex: columnIndex,
        },
        rowSpan: 1,
        columnSpan: 1,
      },
      tableCellProperties: {
        tableCellBackgroundFill: {
          solidFill: {
            color: {
              rgbColor: {red: color[0], green: color[1], blue: color[2]},
            },
          },
        },
      },
      fields: 'tableCellBackgroundFill.solidFill.color',
    },
  });

  documentProperties.setProperty('SLIDES_REQUESTS', JSON.stringify(requests));
}

/**
 * Determines a color based on if a value is a Good, Needs Improvement or Poor
 * range for a given metric.
 *
 * @param {Array} range Array with a low and high threshold for a CWV metric.
 * @param {Number} value Number indicating the metric score.
 * @return {Array} Array of RBG values in decimal form.
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

/**
 * Applies conditional coloring table to the CWV parameter table.
 *
 * @param {!string} deckId Id of the new slide deck that contains the table to
 * be styled.
 */
function colorCWVTable(deckId) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const cwvSlideIndex = documentProperties.getProperty('CWV_SLIDE');
  const cwvSlide = SlidesApp.openById(deckId).getSlides()[cwvSlideIndex];
  const cwvTable = cwvSlide.getTables()[0];

  const lcpColumn = cwvTable.getColumn(2);
  const fidColumn = cwvTable.getColumn(3);
  const clsColumn = cwvTable.getColumn(4);
  const inpColumn = cwvTable.getColumn(5);

  for (let i = 1; i <= 3; i++) {
    const cell = lcpColumn.getCell(i);
    const color = colorForCWV(CWV.LCP, cell.getText().asString());
    buildBackgroundCellColorTableStyleSlidesRequest(
        cwvTable.getObjectId(), i, 2, color);
  }

  for (let i = 1; i <= 3; i++) {
    const cell = fidColumn.getCell(i);
    const color = colorForCWV(CWV.FID, cell.getText().asString());
    buildBackgroundCellColorTableStyleSlidesRequest(
        cwvTable.getObjectId(), i, 3, color);
  }

  for (let i = 1; i <= 3; i++) {
    const cell = clsColumn.getCell(i);
    const color = colorForCWV(CWV.CLS, cell.getText().asString());
    buildBackgroundCellColorTableStyleSlidesRequest(
        cwvTable.getObjectId(), i, 4, color);
  }

  for (let i = 1; i <= 3; i++) {
    const cell = inpColumn.getCell(i);
    const color = colorForCWV(CWV.INP, cell.getText().asString());
    buildBackgroundCellColorTableStyleSlidesRequest(
        cwvTable.getObjectId(), i, 5, color);
  }

  setHyperlinksInCWVTable(cwvTable);
}

/**
 * Sets hyperlinks for URLs in the CWV table and truncates visible URL if
 * necessary.
 *
 * @param {object} tableRef Table class where first column are the CWV URLs.
 */
function setHyperlinksInCWVTable(tableRef) {
  const urlColumn = tableRef.getColumn(0);
  for (let i = 1; i <= 3; i++) {
    const cell = urlColumn.getCell(i);
    const cellTextRange = cell.getText();
    const fullUrl = cellTextRange.asString().replace(/(\r\n|\n|\r)/gm, '');
    if (fullUrl) {
      const visibleUrl = truncateUrl(fullUrl);
      cellTextRange.setText(visibleUrl);
      cellTextRange.getTextStyle().setLinkUrl(fullUrl);
    }
  }
}

/**
 * Truncates a long url to be less than a max length if necessary, ending in
 * ellipsis.
 *
 * @param {string} urlString A potentially long URL to truncate.
 * @return {string} The final string to display.
 */
function truncateUrl(urlString) {
  const maxLineLength = 64;
  const visibleText =
    urlString.length < maxLineLength ? urlString :
    `${urlString.slice(0, maxLineLength)}...`;
  return visibleText;
}

/**
 * Determines & sets the dynamic title of the CWV table slide.
 *
 * @param {!string} newDeckId Id of the new slide deck that has
 * been generated.
 */
function replaceCWVSlideTitle(newDeckId) {
  const presentation = SlidesApp.openById(newDeckId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DICTIONARY_SHEET_NAME'));
  const customStrings = sheet.getDataRange().offset(1, 0).getValues();
  let result = 'Current Core Web Vital status';
  const urlInsights = getUrlInsights(customStrings);
  const hasBadSpeed = urlInsights.lcp > 0;
  const hasBadInteractivity = urlInsights.fid > 0 || urlInsights.inp > 0;
  const hasBadStability = urlInsights.cls > 0;
  const perfectScore = !hasBadSpeed && !hasBadInteractivity && !hasBadStability;
  const worstScore = hasBadSpeed && hasBadInteractivity && hasBadStability;
  if (perfectScore) {
    result = 'Speed, interactivity, and stability metrics are good but could be optimized further';
  } else if (worstScore) {
    result = 'Speed, interactivity, and stability metrics need improvements';
  } else {
    let prefixStr = '';
    if (hasBadSpeed) {
      prefixStr = 'Speed';
    }
    if (hasBadInteractivity) {
      prefixStr += prefixStr.length ? ' and interactivity' : 'Interactivity';
    }
    if (hasBadStability) {
      prefixStr += prefixStr.length ? ' and stability' : 'Stability';
    }
    result = `${prefixStr} metrics need improvements`;
  }
  replaceText(presentation, [['{{current_status_title}}', result]]);
}

/**
 * Gets url insights for determining the dynamic title of the CWV table slide.
 *
 * @param {array} customStrings Nested array of custom strings to be replaced in
 * the final Presentation containing [searchString, replaceString] where
 * searchString is the metric name formatted as {{metric}} (and the metric is 3
 * letters long) and replaceString is the metric score.
 * @return {object} Bad scores object used to determine the dynamic title of the
 * CWV slide.
 */
function getUrlInsights(customStrings) {
  const badScores = {
    lcp: 0,
    cls: 0,
    fid: 0,
    inp: 0,
  };
  for (const row of customStrings) {
    if (!row[0]) break;
    const [searchString, replaceString] = row;
    const metric = searchString?.slice(2, 5);
    if (badScores[metric] !== undefined) {
      if (metric === 'lcp' && replaceString) {
        badScores[metric] += replaceString >= CWV.LCP[0];
      }
      if (metric === 'cls' && replaceString) {
        badScores[metric] += replaceString >= CWV.CLS[0];
      }
      if (metric === 'fid' && replaceString) {
        badScores[metric] += replaceString >= CWV.FID[0];
      }
      if (metric === 'inp' && replaceString) {
        badScores[metric] += replaceString >= CWV.INP[0];
      }
    }
  }
  return badScores;
}
