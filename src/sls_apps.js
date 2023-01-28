/* exported createDeckFromDatasource */
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
 * @fileoverview Functions required by the Apps version of Slide Starter
 * to automate the creation of slides for a MetricsFlow audits, ran by the mApps
 * team.
 *
 * This file requires the following functions to be "imported" from the core
 * UX starter script to function properly:
 *
 * - loadConfiguration
 *   Populates Document Properties with the parameters retrieved from the
 *   configuration sheet on the trix
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
 * - createBaseDeck
 *   Creates a slide deck based on a base template
 *
 * Slide Starter (Based on UX Starter V5 - Last edit 11/02/22)
 */

// TODO: Move into configuration sheet.
const COLUMN_WIDTH = [50, 50, 75, 250, 75, 180];
const MAX_PAGE_HEIGHT = 175;
const LINE_HEIGHT = 10;
const ROW_HEIGHT = 20;

/**
 * Creates a presentation based on a set of recommendations included in a
 * spreadsheet. For this, it creates a copy of a base deck, it retrieves the
 * templates* for each type of slide (header and table), and creates as many
 * sections in the deck as there were indicated in the configuration sheet with
 * a comma-separated list.
 *
 * Since some of these operations require the advanced Slides API service for
 * formatting, these operation requests are applied after all slides have been
 * created. The requests object, which is stored globally in Document Properties
 * (after being flattened) is reinitialized in this function so that no
 * properties from previous executions are being carried over.
 *
 * *While the templates are not being used in a few levels, given that the
 * process to retrieve them is computationally intensive, it's done at this
 * level to avoid repeating that operation.
 */
function createDeckFromDatasource() {
  loadConfiguration();
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('SLIDES_REQUESTS', JSON.stringify([]));
  const newDeckId = createBaseDeck();
  const recommendationSlideLayout =
      getTemplateLayout(newDeckId, 'TABLE_LAYOUT_NAME');
  const headerSlideLayout = getTemplateLayout(newDeckId, 'HEADER_LAYOUT_NAME');

  const auditSheets =
      documentProperties.getProperty('DATA_SOURCE_SHEET').split(',');

  for (const sheetName of auditSheets) {
    createSlideSection(
        sheetName.trim(), newDeckId, headerSlideLayout,
        recommendationSlideLayout);
  }
  const resource = {
    requests: JSON.parse(documentProperties.getProperty('SLIDES_REQUESTS')),
  };
  Slides.Presentations.batchUpdate(resource, newDeckId);
}

/**
 * Handles the creation of a logical section of slides for MetricFlow audit deck
 * which are comprised of a header slide (based on the section name) and one or
 * more pages containing a table with the results of the audit.
 *
 * @param {string} sectionName Name of the section of the audit
 * @param {string} deckId Object identifier for the slide deck
 * @param {!Layout} headerSlideLayout Layout object relative to the header slide
 * @param {!Layout} recommendationSlideLayout Layout object for table slides
 */
function createSlideSection(
    sectionName, deckId, headerSlideLayout, recommendationSlideLayout) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
      'Creating slides for section: \n' + sectionName);
  const documentProperties = PropertiesService.getDocumentProperties();
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(sectionName);
  const startingRow = documentProperties.getProperty('STARTING_ROW');
  const values = spreadsheet
      .getRange(
          startingRow, 1, spreadsheet.getLastRow() - startingRow,
          spreadsheet.getLastColumn())
      .getValues();
  if (values.length === 1) {
    return;
  }

  createHeaderSlide(deckId, headerSlideLayout, sectionName);
  createPaginatedTableSlides(
      deckId, recommendationSlideLayout, sectionName, values);
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
 * Dynamically generates a paginated table based on a possible height overflow
 * given the values of the rows of the table incoming from the audit spreadsheet
 *
 * Given that the number of elements per page and their height are both variable
 * this function considers whether there would be an overflow if another row was
 * to be included, and marks the page as full, which then triggers the creation
 * of a table slide, emptying the page.
 *
 * Handling elements after possible overflow conditions are met is as follows:
 * - If the page fills due to height overflow, the row is returned to values
 * - If the page fills due to no more elements in values, the last row is pushed
 *   into the page
 *
 * Finally, if there no elements left on the original values array
 * but the page isn't full, it will create the last slide.
 *
 * @param {string} deckId Object identifier for the slide deck
 * @param {!Layout} layout Layout object relative to the header slide
 * @param {string} sectionName Name of the section of the audit
 * @param {!Array<!Array<string>>} values Multidimensional array containing the
 *     values of the table as retrieved from the spreadsheet
 */
function createPaginatedTableSlides(deckId, layout, sectionName, values) {
  const headerRow = values.shift();
  let page = [];
  let currentPageHeight = 0;
  let pageFull = false;
  while (values.length > 0) {
    const currentRow = values.shift();
    currentPageHeight += calculateRowHeight(currentRow);
    if (currentPageHeight > MAX_PAGE_HEIGHT) {
      // TODO: Address possible edge case where single row overflows
      values.unshift(currentRow);
      pageFull = page.length > 0;
    } else if (values.length === 0) {
      page.push(currentRow);
      pageFull = true;
    } else {
      page.push(currentRow);
    }
    if (pageFull) {
      page.unshift(headerRow);
      createTableSlide(deckId, layout, sectionName, page);
      pageFull = false;
      page = [];
      currentPageHeight = 0;
    }
  }
}

/**
 * Creates a table slide within a specified deck and appends it, then filling it
 * with the values corresponding to the rows passed as parameter. If the cell
 * isn't empty, it passes the contents to a function that will apply style to
 * the textrange.
 *
 * @param {string} deckId Object identifier for the slide deck
 * @param {!Layout} layout Layout object relative to the header slide
 * @param {string} title Name of the section of the audit
 * @param {!Array<!Array<string>>} values Multidimensional array containing the
 *     values of the table as retrieved from the spreadsheet
 */
function createTableSlide(deckId, layout, title, values) {
  const deck = SlidesApp.openById(deckId);
  const slide = deck.appendSlide(layout);
  const titlePlaceholder =
      slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  const titleRange = titlePlaceholder.asShape().getText();
  titleRange.setText(title);
  const tableShape = retrieveShape(slide, 'table_shape');
  const table = slide.insertTable(
      values.length, values[0].length, tableShape.getLeft(),
      tableShape.getTop(), tableShape.getWidth(), values.length * ROW_HEIGHT);
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const isHeaderRow = i === 0;
    for (let j = 0; j < row.length; j++) {
      const cellText = table.getCell(i, j).getText();
      cellText.setText(row[j]);
      if (cellText.asString().trim() !== '') {
        applyTextStyle(cellText, isHeaderRow);
      }
    }
  }

  deck.saveAndClose();
  buildTableStyleSlidesRequest(slide.getTables()[0].getObjectId());
}

/**
 * Applies a style to a text range based on whether this text is in the header
 * or not.
 *
 * @param {!TextRange} textRange Text range retrieved from the table cell
 * @param {boolean} isHeader flag to determine whether header styling applies
 */
function applyTextStyle(textRange, isHeader) {
  const style = textRange.getTextStyle();
  // TODO: Move configurable params to configuration sheet.
  if (isHeader) {
    style.setFontSize(9);
    style.setFontFamily('Roboto');
    style.setForegroundColor('#FFFFFF');
    style.setBold(true);
  } else {
    style.setFontSize(9);
    style.setFontFamily('Roboto');
    const stringContent = textRange.asString().trim();
    if (stringContent === 'Yes') {
      style.setForegroundColor('#1E8E3E');
    } else if (stringContent === 'No') {
      style.setForegroundColor('#A50E0E');
    }
  }
}


/**
 * Builds a SlidesAPI request to handle the table formatting properties that
 * are not accessible via the SlidesAPI service, such as column width.
 * These requests are retrieved and stored from the document properties after
 * being flattened as JSON.
 *
 * @param {string} tableId String that identifies the table to modify
 */
function buildTableStyleSlidesRequest(tableId) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const requests =
      JSON.parse(documentProperties.getProperty('SLIDES_REQUESTS'));

  // TODO: Make column width dynamic from configuration sheet.
  for (let i = 0; i < COLUMN_WIDTH.length; i++) {
    requests.push({
      updateTableColumnProperties: {
        tableColumnProperties: {
          columnWidth: {
            magnitude: COLUMN_WIDTH[i],
            unit: 'PT',
          },
        },
        columnIndices: [i],
        objectId: tableId,
        fields: 'columnWidth',
      },
    });
  }

  // TODO: Make color dynamic from configuration sheet.
  requests.push({
    updateTableCellProperties: {
      objectId: tableId,
      tableRange: {
        location: {
          rowIndex: 0,
          columnIndex: 0,
        },
        rowSpan: 1,
        columnSpan: 6,
      },
      tableCellProperties: {
        tableCellBackgroundFill: {
          solidFill: {
            color: {
              rgbColor: {red: 0.26, green: 0.52, blue: 0.96},
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
 * This function calculates the height of a row in points based on the maximum
 * possible height of each individual cells.
 *
 * @param {!Array<string>} row Array with the contents of each cell in a row
 * @return {number} Maximum possible height of a row, in PT
 */
function calculateRowHeight(row) {
  let currentRowHeight = 0;
  for (let j = 0; j < row.length; j++) {
    const currentColHeight =
        calculateCellRowHeight(row[j].length, j, LINE_HEIGHT);
    if (currentColHeight > currentRowHeight) {
      currentRowHeight = currentColHeight;
    }
  }
  return currentRowHeight;
}

/**
 *
 * @param {number} lengthText Length of the text contained in the cell
 * @param {number} columnIndex 0-based index of the column within the row
 * @param {number} fontSize Size of the font, in PT, including padding
 * @return {number} Calculated cell height in PT
 */
function calculateCellRowHeight(lengthText, columnIndex, fontSize) {
  return Math.ceil(lengthText * fontSize / COLUMN_WIDTH[columnIndex]) *
      fontSize;
}
