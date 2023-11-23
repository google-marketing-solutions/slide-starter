function recordDeckCreated(newDeckId = '') {
  const telemetryArray = [];
  const sheetId = SpreadsheetApp.getActive().getId();
  const documentProperties = PropertiesService.getDocumentProperties();
  const telemetryString = documentProperties.getProperty('TELEMETRY_DATA');
  const clientInfoArray = telemetryString.split(';');
  const today = new Date().toISOString().slice(0, 10);

  telemetryArray.push(sheetId, ...clientInfoArray, newDeckId, today);
  const telemetryTargetId = documentProperties.getProperty('TELEMETRY_TARGET_ID');
  const targetSheet = SpreadsheetApp.openById(telemetryTargetId);
  targetSheet.getSheetByName('Data');
  targetSheet.appendRow(telemetryArray);
}
