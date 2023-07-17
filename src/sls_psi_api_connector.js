/**
 * @typedef {{
*            error: {message: string}} |
*            {lighthouseResult: !Object, loadingExperience: !Object}}
*/
let PsiResult;

// Global variables used throughout the solution.
const CLIENT_REQUESTS_SHEET = "Performance";
const RESULTS_TAB = "Performance Results";

/**
* Reads PSI API Key from the Sheet.
*
* If no string is found in the appropriate cell, an alert is shown in the
* sheet.
*
* @return {string} the API Key to use with PSI.
*/
function getPsiApiKey() {
 const documentProperties = PropertiesService.getDocumentProperties();
 /** @type {string} */ documentProperties.getProperty('PSI_API_KEY');
 if (!key.trim()) {
   SpreadsheetApp.getUi().alert("Please enter your API Key");
   throw new Error("The PSI API key must be set to use this tool.");
 }
 return key;
}

/**
* Triggers the tests and outputs the results to the Sheet.
*/
function runPSITests() {
 const urlSettings = getURLSettings();
 const responses = submitTests(urlSettings);
 const sheet = SpreadsheetApp.getActive().getSheetByName(RESULTS_TAB);
 const today = new Date().toISOString().slice(0, 10);
 const responseMap = createResultsMap()
 // There should be one response for each row of urlSettings.
 for (let i = 0; i < responses.length; i++) {
   const url = urlSettings[i][0]; // A
   const label = urlSettings[i][1]; // B
   const device = urlSettings[i][2]; // C

   const content = /** @type {!PsiResult} */ (
     JSON.parse(responses[i].getContentText())
   );
   if (content.error) {
     sheet.appendRow([url, label, device]);
     const note =
       `${content.error.message}\n\n` +
       "If this error persists, investigate the cause by running the " +
       "URL manually via " +
       "https://developers.google.com/speed/pagespeed/insights/";
     addNote(note, "#fdf6f6"); // light red background
   } else {
     const results = parseResults(content, responseMap);
     let cruxDataType = 'PAGE';
     if (!results.crux_data) {
      cruxDataType = 'NONE';
    } else if (results.origin_fallback) {
      cruxDataType = 'ORIGIN';
    }
    const resultsData = [url, label, device, today, cruxDataType, ...results.data];
    sheet.appendRow(resultsData); 
   }
 }

 const results = SpreadsheetApp.getActive().getSheetByName(RESULTS_TAB);
}

/**
* Reads and then deletes rows from the from queue.
*
* @return {!Array<!Array<(string | number)>>} An array with all the settings
*     for each URL.
*/
function getURLSettings() {
 const sheet = SpreadsheetApp.getActive().getSheetByName(
   CLIENT_REQUESTS_SHEET
 );
 const lastColumn = sheet.getLastColumn();
 let lastRow = sheet.getLastRow() - 1;
 const range = sheet.getRange(2, 1, lastRow, lastColumn);
 const settings = /** @type {!Array<!Array<(string | number)>>} */ (
   range.getValues()
 );
 return settings;
}

/**
* Builds the fetch URLs for PSI and submits them in parallel.
*
* The format of a request to PSI is documented here:
* https://developers.google.com/speed/docs/insights/v5/reference/pagespeedapi/runpagespeed#request
*
* @param {!Array<!Array<(string | number)>>} settings The URL settings for
*     all
*    tests.
* @return {!Array<!GoogleAppsScript.URL_Fetch.HTTPResponse>} All the responses
*     from PSI.
*/
function submitTests(settings) {
 const key = getPsiApiKey();
 const categories = "&category=BEST_PRACTICES" + "&category=PERFORMANCE";
 const serverURLs = settings.map(([url, unused, device]) => ({
   url: `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?${categories}&strategy=${device}&url=${url}&key=${key}`,
   muteHttpExceptions: true,
 }));
 const responses = UrlFetchApp.fetchAll(serverURLs);
 return responses;
}

/**
* Parses the response from PSI and prepares it for the sheet.
*
* The format of the response from PSI is documented here:
* https://developers.google.com/speed/docs/insights/v5/reference/pagespeedapi/runpagespeed#response
*
* @param {!PsiResult} content The
*     lighthouseResult object returned from PSI to parse.
* @return {{data: !Array<number | string>, crux_data: boolean, origin_fallback:
*     boolean}} Post-processed data as an array and flags for how the CrUX data
*     was reported.
*/
function parseResults(content, responseMap) {
 const allResults = {
   data: [],
   crux_data: false,
   origin_fallback: false,
 };
 const { lighthouseResult, loadingExperience } = content;
 const version = lighthouseResult["lighthouseVersion"];
 const screenshot = lighthouseResult["audits"]["final-screenshot"]["details"]["data"];
 const categories = [];
 responseMap.get("categories").forEach((category) => {
   const score = lighthouseResult["categories"][category]["score"] * 100;
   categories.push(score);
 });
 const metrics = [];
 responseMap.get("metrics").forEach((metric) => {
   const value = lighthouseResult["audits"][metric]["numericValue"];
   metrics.push(value);
 });
 const auditKeys = Object.keys(lighthouseResult["audits"]);
 const failedAuditNames = auditKeys.filter((auditName) => {
  const score = lighthouseResult["audits"][auditName].score;
  if (score < 1 && score != null) {
    return auditName;
  }
 });

 const crux = [];
 if (loadingExperience["metrics"]) {
   allResults.crux_data = true;
   crux.push(loadingExperience["overall_category"]);
   responseMap.get("crux").forEach((metricName) => {
     if (loadingExperience["metrics"][metricName]) {
       const metric = loadingExperience["metrics"][metricName];
       let percentile = metric["percentile"];
       if (metricName === "CUMULATIVE_LAYOUT_SHIFT_SCORE") {
         percentile = percentile / 100;
       }
       crux.push(percentile);
     } else {
       crux.push([, , , , ,]); // filler for the sheet if the metric isn't there
     }
   });
   // If there's insufficient field data for the page, the API responds with
   // origin-level field data and origin_fallback = true.
   if (loadingExperience["origin_fallback"]) {
     allResults.origin_fallback = true;
   }
 }
 allResults.data = [screenshot, ...crux, ...categories, ...metrics, failedAuditNames.toString(), version];
 return allResults;
}

/**
* Attaches an info note to the current last row of the sheet.
*
* @param {string} note The note to add.
* @param {?string=} formatColor The background color of the note in rgb
*     hex. The default null value leaves the color as is.
*/
function addNote(note, formatColor = null) {
 const sheet = SpreadsheetApp.getActive().getSheetByName(RESULTS_TAB);
 const lastRow = sheet.getLastRow();
 sheet.getRange(`${lastRow}:${lastRow}`).setBackground(formatColor);
 sheet.getRange(`D${lastRow}`).setNote(note);
}

/**
* Creates a results map for the given row of the URL settings array.
// These are used to index the objects returned from PSI,
// which is why they are named as they are. The order they are defined here
// is also the order they are inserted into the sheet per parseResults.
*
* @return {!Map<string, !Map<string, number>>} The budget values in an object.
*/
function createResultsMap() {
  const categories = ["performance"];
  const crux = [
    "FIRST_CONTENTFUL_PAINT_MS",
    "LARGEST_CONTENTFUL_PAINT_MS",
    "FIRST_INPUT_DELAY_MS",
    "CUMULATIVE_LAYOUT_SHIFT_SCORE",
    "INTERACTION_TO_NEXT_PAINT",
  ];
  const metrics = [
    "server-response-time",
    "first-contentful-paint",
    "largest-contentful-paint",
    "total-blocking-time",
    "cumulative-layout-shift"
  ];
  const assets = [
    "total",
    "script",
    "image",
    "stylesheet",
    "document",
    "font",
    "other",
    "media",
    "third-party",
  ];


  const requiredResultsMap = new Map();
  requiredResultsMap.set("categories", categories);
  requiredResultsMap.set("metrics", metrics);
  requiredResultsMap.set("assets", assets);
  requiredResultsMap.set("crux", crux);
  return requiredResultsMap;
}

//  audits = [];
//  failedAuditNames.forEach((failedAuditName) => {
//   audits.push(lighthouseResult["audits"][failedAuditName].title);
//  });
 // const resources =
 //   lighthouseResult["audits"]["resource-summary"]["details"]["items"];
 // const assetsObject = Object.fromEntries(
 //   resources.map((resource) => [resource.resourceType, resource])
 // );
 // const assets = [];
 // responseMap.get("assets").forEach((assetType) => {
 //   const transferSize = assetsObject[assetType]["transferSize"] / 1024;
 //   assets.push(transferSize, assetsObject[assetType]["requestCount"]);
 // });
 