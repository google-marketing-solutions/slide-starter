/**
 * @fileoverview Description of this file.
 */

const CWV = {
  LCP: [2500, 4000],
  FID: [100, 300],
  CLS: [0.1, 0.25],
};
const NUM_PROPERTIES = 16;

/**
 * Parses the fields contained on the incoming row from the spreadsheet into
 * some specific information fields, and then creates the slide using GAS.
 *
 * @param {!Presentation} deck Id of the generated deck that will contain the
 *     recos
 * @param {!Presentation} insightDeck Reference to the generated deck
 * @param {!Layout} recommendationSlideLayout The template layout
 * @param {!Array<string>} row Array of strings with information from the
 *     spreadsheet
 */
function parseFieldsAndCreateSlide(
    deck, insightDeck, recommendationSlideLayout, row) {
  const criteriaNameIndex =
      documentProperties.getProperty('RECOMMENDATIONS_CRITERIA_NAME_ROW') - 1;
  const criteriaAppliesIndex =
      documentProperties.getProperty('RECOMMENDATIONS_APPLIES_ROW') - 1;
  const criteriaProblemStatementIndex =
      documentProperties.getProperty('RECOMMENDATIONS_PROBLEM_STATEMENT_ROW') -
      1;
  const criteriaSolutionStatementIndex =
      documentProperties.getProperty('RECOMMENDATIONS_SOLUTION_STATEMENT_ROW') -
      1;
  const criteriaInsightSlidesIndex =
      documentProperties.getProperty('RECOMMENDATIONS_INSIGHTS_ROW') - 1;

  const criteria = row[criteriaNameIndex];
  const applicable =
      `Applies for: ${row[criteriaAppliesIndex].split(',').join(',')}`;
  const description = row[criteriaProblemStatementIndex];
  const learnMore = row[criteriaSolutionStatementIndex];
  const insights = row[criteriaInsightSlidesIndex].split(',');

  createRecommendationSlideGAS(
      deck, recommendationSlideLayout, criteria, applicable, description,
      learnMore, insights);
  if (insights.length > 0) {
    appendInsightSlides(deck, insightDeck, insights);
  }
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

  documentProperties.setProperty('SLIDES_REQUESTS', JSON.stringify([]));
  colorCWVTable(newDeckId);
  const resource = {
    requests: JSON.parse(documentProperties.getProperty('SLIDES_REQUESTS')),
  };
  Slides.Presentations.batchUpdate(resource, newDeckId);
}

/**
 * Creates the slides programmatically using the SlidesApp from AppScript:
 * It first creates a new slide with the specified layout, it populates the
 * placeholders with
 *
 * @param {string} deck Id of the generated deck that will contain the recos
 * @param {!Layout} recommendationSlideLayout The template layout
 * @param {string} criteria The name of the criteria used as title
 * @param {string} applicable A list of pages where this criteria is applicable
 * @param {string} description The description of the failing criteria
 * @param {string} learnMore The URL of the page with extended information
 */
function createRecommendationSlideGAS(
    deck, recommendationSlideLayout, criteria, applicable, description,
    learnMore) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
      'Creating slide for criteria: ' + criteria);
  const slide = deck.appendSlide(recommendationSlideLayout);

  const titlePlaceholder =
      slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  // const subtitlePlaceholder =
  //     slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
  const bodyPlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);

  const titleRange = titlePlaceholder.asShape().getText();
  titleRange.setText(criteria);

  // const subtitleRange = subtitlePlaceholder.asShape().getText();
  // subtitleRange.setText(applicable);

  const bodyRange = bodyPlaceholder.asShape().getText();
  bodyRange.setText(description);

  const learnMoreShape = retrieveShape(slide, 'learn_more');
  learnMoreShape.setLinkUrl(learnMore);
  slide
      .insertTextBox(
          learnMore, learnMoreShape.getLeft(), learnMoreShape.getTop(),
          learnMoreShape.getWidth(), learnMoreShape.getHeight())
      .setLinkUrl(learnMore);
}

/**
 * Builds a SlidesAPI request to handle the table formatting properties that
 * are not accessible via the SlidesAPI service, such as column width.
 * These requests are retrieved and stored from the document properties after
 * being flattened as JSON.
 *
 * @param {string} tableId String that identifies the table to modify
 */
function buildBackgroundCellColorTableStyleSlidesRequest(
    tableId, rowIndex, columnIndex, color) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const requests = JSON.parse(documentProperties.getProperty('SLIDES_REQUESTS'));

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


function colorForCWV(cwv, value) {
  const lowThreshold = cwv[0];
  const highThreshold = cwv[1];

  if (value <= lowThreshold) {
    return [.04, .80, .41];
  } else if (value < highThreshold) {
    return [1, 0.64, 0];
  } else {
    return [1, 0.30, 0.25];
  }
}

function colorCWVTable(deckId) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const cwvSlideIndex = documentProperties.getProperty('CWV_SLIDE');
  const cwvSlide = SlidesApp.openById(deckId).getSlides()[cwvSlideIndex];
  const cwvTable = cwvSlide.getTables()[0];

  const lcpColumn = cwvTable.getColumn(2);
  const fidColumn = cwvTable.getColumn(3);
  const clsColumn = cwvTable.getColumn(4);

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
}
