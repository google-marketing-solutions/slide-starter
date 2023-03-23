function buildReadinessAnalysis(spreadsheet, values, chartSheetName) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const policyNamesListString = documentProperties.getProperty('CATEGORY_NAMES_LIST');
  const policyNamesList = policyNamesListString.split(',').map( item => item.trim());
  const policyColumnIndex = documentProperties.getProperty('POLICY_MAPPING_COLUMN') - 1;
  const policyValuesList = new Array(policyNamesList.length).fill(0);
  const policyTotalList = new Array(policyNamesList.length).fill(0);

  for (let row of values) {
    if (!row[policyColumnIndex]) {
      continue;
    }
    let rowPolicyArray = row[policyColumnIndex].split(',').map( item => item.trim());
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

  const partialValuesRange = "'" + chartSheetName + "'!" + documentProperties.getProperty('PARTIAL_RESULTS_RANGE');
  const totalValuesRange = "'" + chartSheetName + "'!" + documentProperties.getProperty('TOTAL_RESULTS_RANGE');
  SpreadsheetApp.getActive().getRangeByName(partialValuesRange).setValues([policyValuesList]);
  SpreadsheetApp.getActive().getRangeByName(totalValuesRange).setValues([policyTotalList]);
}