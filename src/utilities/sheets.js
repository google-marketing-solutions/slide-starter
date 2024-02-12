/**
 * @license
 * Copyright 2024 Google LLC
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
 * Retrieves the active spreadsheet,removes the current filter if it exists,
 * applies new filter based on criteria, and sorts by a specified column in the
 * trix.
 *
 * @param {!Sheet=} sheet - The sheet to apply the filter and sort to. Defaults
 *     to the active sheet.
 */
function filterAndSortData(sheet = undefined) {
  SpreadsheetApp.getActiveSpreadsheet().toast('Filtering and sorting');
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
 * Below are the exports required for the linter.
 * This is necessary because AppsScript doesn't support modules.
 */
/* exported filterAndSortData */
