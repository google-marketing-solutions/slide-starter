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
 * Google AppScript File
 * @fileoverview Includes the core shared functions between the different
 * implementations of Slide Starter for the procedural generation of slide
 * decks.
 * - loadConfiguration
 *   Populates Document Properties with the parameters retrieved from the
 *   configuration sheet on the trix
 *
 * - createBaseDeck
 *   Creates a slide deck based on a base template
 *
 * - customDataInjection
 *   Finds and replaces all placeholder strings within a slide deck
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
 * 23/05/23
 */

// --- Katalyst loops

/**
 * Creates a new Slides deck based on the data sources specified in the document
 * properties. Uses the specified base deck as a template, and applies custom
 * styling to the new deck.
 */
function createDeckFromDatasources() {
  loadConfiguration();

  const newDeckId = createBaseDeck();

  const datasourceString = documentProperties.getProperty('DATA_SOURCE_SHEET');
  const datasourcesArray =
      datasourceString.split(',').map((item) => item.trim());

  const sectionLayoutName =
      documentProperties.getProperty('SECTION_LAYOUT_NAME');
  let sectionLayout;
  if (sectionLayoutName && sectionLayoutName.length > 0) {
    sectionLayout = getTemplateLayout(newDeckId, sectionLayoutName);
  }

  for (const datasource of datasourcesArray) {
    if (sectionLayout) {
      createSlideWithTitle(deckId, sectionLayout, datasource);
    }
    prepareDependenciesAndCreateSlides(datasource, newDeckId);
  }

  const dictionarySheetName =
      documentProperties.getProperty('DICTIONARY_SHEET_NAME');
  if (dictionarySheetName && dictionarySheetName.length > 0) {
    customDataInjection(newDeckId);
  }

  applyCustomStyle(newDeckId);
}

/**
 * Prepares the dependencies and creates slides for the given datasource.
 *
 * @param {string} datasource The name of the datasource.
 * @param {string} newDeckId The ID of the new deck to create slides in.
 */
function prepareDependenciesAndCreateSlides(datasource, newDeckId) {
  const datasourceConfiguration =
      '\'Configuration_' + datasource + '\'!PROPERTIES';
  loadConfiguration(datasourceConfiguration);

  const deck = SlidesApp.openById(newDeckId);
  const recommendationSlideLayout = getTemplateLayout(newDeckId);

  let insightDeck;
  const insightsDeckId = documentProperties.getProperty('INSIGHTS_DECK_ID');
  if (insightsDeckId && insightsDeckId.length > 0) {
    insightDeck = SlidesApp.openById(insightsDeckId);
  }

  createSlidesForDatasource(deck, insightDeck, recommendationSlideLayout);
}

/**
 * If a custom function is specified in the configuration tab, the custom
 * function is called instead. Otherwise t creates either a single slide or a
 * collection slide based on the check in config.
 *
 * @param {!Presentation} deck - The Slides deck where the new slide(s) will be
 *     created.
 * @param {!Presentation} insightDeck - Extra deck to pull insight slides,
 *     retrieved only once.
 * @param {!Layout} slideLayout - The slide layout to use for the new slide(s).
 *
 */
function createSlidesForDatasource(deck, insightDeck, slideLayout) {
  const customFunctionName = documentProperties.getProperty('CUSTOM_FUNCTION');
  if (customFunctionName && customFunctionName.length > 0) {
    getFunctionByName(customFunctionName)(deck, insightDeck, slideLayout);
  } else {
    const isSingleSlide =
        documentProperties.getProperty('SINGLE_VALUE') == 'TRUE';
    if (isSingleSlide) {
      createSingleSlide(deck, insightDeck, slideLayout);
    } else {
      createCollectionSlide(deck, insightDeck, slideLayout);
    }
  }
}

/**
 * Creates a collection slide based on data from a Google Sheets data source
 * using the specified deck, insight deck, and slide layout. Filters and sorts
 * the data, and creates a slide for each row that passes the filter criteria.
 *
 * @param {!Presentation} deck - The Slides deck where the new slide(s) will be
 *     created.
 * @param {!Presentation} insightDeck - The Slides deck where the insight slide
 *     will be created (if applicable).
 * @param {!Layout} slideLayout - The slide layout to use for the new slide(s).
 *
 */
function createCollectionSlide(deck, insightDeck, slideLayout) {
  // Execute the pre-collection creation hook
  const preCollectionFunction =
      documentProperties.getProperty('PRE_COLLECTION_FUNCTION');
  if (preCollectionFunction && preCollectionFunction.length > 0) {
    const preCollectionFunctionArgsRaw =
        documentProperties.getProperty('PRE_COLLECTION_FUNCTION_ARGS');
    let preCollectionFunctionArgs;
    if (preCollectionFunctionArgsRaw &&
        preCollectionFunctionArgsRaw.length > 0) {
      preCollectionFunctionArgs = preCollectionFunctionArgsRaw.split(',');
    } else {
      preCollectionFunctionArgs = [];
    }
    getFunctionByName(preCollectionFunction)(
        deck, ...preCollectionFunctionArgs);
  }
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));
  filterAndSortData();
  const values = spreadsheet.getFilter().getRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (spreadsheet.isRowHiddenByFilter(i + 1)) {
      continue;
    }
    const row = values[i];
    if (shouldCreateCollectionSlide()) {
      parseFieldsAndCreateCollectionSlide(deck, slideLayout, row);
    }
    addInsightSlides(deck, insightDeck, row);
  }
}

/**
 * Creates a single slide in a Google Slides presentation.
 * The function first gets the spreadsheet that contains the data for the slide.
 * Then, it creates a new slide in the presentation using the specified layout.
 * If there is a master slide, the function removes it. Next, the function
 * fetches the title, subtitle, and body text for the slide from the
 * spreadsheet. It then sets the title, subtitle, and body text for the slide.
 * Finally, the function fetches the image shapes and ranges for the slide from
 * the spreadsheet. If there are image shapes and ranges, the function adds the
 * images to the slide.
 *
 *
 * @param {!SlidesApp.Presentation} deck The presentation to add the slide to.
 * @param {!SlidesApp.InsightDeck} insightDeck The insight deck that contains
 *     the data for the slide.
 * @param {!SlidesApp.SlideLayout} slideLayout The layout to use for the slide.
 *
 * @return {void}
 */
function createSingleSlide(deck, insightDeck, slideLayout) {
  const spreadsheet = SpreadsheetApp.getActive().getSheetByName(
      documentProperties.getProperty('DATA_SOURCE_SHEET'));

  const slide = deck.appendSlide(slideLayout);
  if (deck.getMasters().length > 1) {
    deck.getMasters()[deck.getMasters().length - 1].remove();
  }

  // Fetch fields
  const titleRange = documentProperties.getProperty('TITLE_RANGE');
  if (titleRange && titleRange.length > 0) {
    const title = spreadsheet.getRange(titleRange).getValue();
    const slideTitlePlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    const slideTitle = slideTitlePlaceholder.asShape().getText();
    slideTitle.setText(title);
  }


  const subtitleRange = documentProperties.getProperty('SUBTITLE_RANGE');
  if (subtitleRange && subtitleRange.length > 0) {
    const subtitle = spreadsheet.getRange(subtitleRange).getValue();
    const slideSubtitlePlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    const slideSubtitle = slideSubtitlePlaceholder.asShape().getText();
    slideSubtitle.setText(subtitle);
  }

  const bodyRange = documentProperties.getProperty('BODY_RANGE');
  if (bodyRange && bodyRange.length > 0) {
    const body = spreadsheet.getRange(bodyRange).getValue();
    const slideBodyPlaceholder =
        slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    const slideBody = slideBodyPlaceholder.asShape().getText();
    slideBody.setText(body);
  }

  const imageShapesArray = documentProperties.getProperty('IMAGE_SHAPES')
      ?.split(',')
      .map((item) => item.trim());
  const imageRangesArray = documentProperties.getProperty('IMAGE_RANGES')
      ?.split(',')
      .map((item) => item.trim());

  if (imageShapesArray && imageShapesArray.length > 0) {
    for (let i = 0; i < imageShapesArray.length; i++) {
      const shapeId = imageShapesArray[i];
      const range = imageRangesArray[i];

      if (shapeId && range) {
        const imageShape = retrieveShape(slide, shapeId);
        const rawImageValue = spreadsheet.getRange(range).getValue();
        const imageValue = getImageBlobFromFolder(rawImageValue, column);
        slide.insertImage(
            imageValue, imageShape.getLeft(), imageShape.getTop(),
            imageShape.getWidth(), imageShape.getHeight());
      }
    }
  }
}


/**
 * Creates a collection slide based on a slide layout and data from a specified
 * row in a Google Sheet.
 * @param {!GoogleAppsScript.Slides.Presentation} deck - The slide deck to add
 *     the slide to.
 * @param {!GoogleAppsScript.Slides.Layout} slideLayout - The layout to use for
 *     the slide.
 * @param {!Array<string>} row Array of strings with information from the
 *     spreadsheet
 */
function parseFieldsAndCreateCollectionSlide(deck, slideLayout, row) {
  // Execute the pre-slide creation hook
  const preSlideFunction = documentProperties.getProperty('PRE_SLIDE_FUNCTION');
  if (preSlideFunction && preSlideFunction.length > 0) {
    getFunctionByName(preSlideFunction)(deck, slideLayout, row);
  }

  // Create slide
  const slide = deck.appendSlide(slideLayout);
  if (deck.getMasters().length > 1) {
    deck.getMasters()[deck.getMasters().length - 1].remove();
  }

  // Add title
  const titleColumn = documentProperties.getProperty('TITLE_COLUMN');
  if (titleColumn && titleColumn.length > 0) {
    const titleText = row[titleColumn - 1];
    addTextToPlaceholder(slide, SlidesApp.PlaceholderType.TITLE, titleText, '');
  }

  // Add subtitle
  const subtitleColumn = documentProperties.getProperty('SUBTITLE_COLUMN');
  if (subtitleColumn && subtitleColumn.length > 0) {
    const subtitleText = row[subtitleColumn - 1];
    addTextToPlaceholder(
        slide, SlidesApp.PlaceholderType.SUBTITLE, subtitleText, '');
  }

  // Add body
  const bodyColumn = documentProperties.getProperty('BODY_COLUMN');
  if (bodyColumn && bodyColumn.length > 0) {
    const bodyText = row[bodyColumn - 1];
    addTextToPlaceholder(slide, SlidesApp.PlaceholderType.BODY, bodyText, '');
  }

  // Add images
  const imageShapesArray = documentProperties.getProperty('IMAGE_SHAPES')
      ?.split(',')
      .map((item) => item.trim());
  const imageColumnsArray = documentProperties.getProperty('IMAGE_COLUMNS')
      ?.split(',')
      .map((item) => item.trim());

  if (imageShapesArray && imageShapesArray.length > 0) {
    for (let i = 0; i < imageShapesArray.length; i++) {
      const shapeId = imageShapesArray[i];
      const column = imageColumnsArray[i];
      if (shapeId && column) {
        const imageShape = retrieveShape(slide, shapeId);
        const imageValue = getImageBlobFromFolder(row[column - 1], column);
        slide.insertImage(
            imageValue, imageShape.getLeft(), imageShape.getTop(),
            imageShape.getWidth(), imageShape.getHeight());
      }
    }
  }

  // Add other text fields
  const textShapesArray = documentProperties.getProperty('TEXT_SHAPES')
      ?.split(',')
      .map((item) => item.trim());
  const textColumnsArray = (documentProperties.getProperty('TEXT_COLUMNS'))
      ?.split(',')
      .map((item) => item.trim());

  if (textShapesArray && textColumnsArray.length > 0) {
    for (let i = 0; i < textShapesArray.length; i++) {
      const shapeId = textShapesArray[i];
      const column = textColumnsArray[i];

      if (shapeId && column) {
        const textShape = retrieveShape(slide, shapeId);
        const textValue = row[column - 1];
        if (textValue) {
          slide.insertTextBox(
              textValue, textShape.getLeft(), textShape.getTop(),
              textShape.getWidth(), textShape.getHeight());
        }
      }
    }
  }

  const postSlideFunction =
      documentProperties.getProperty('POST_SLIDE_FUNCTION');
  if (postSlideFunction && postSlideFunction.length > 0) {
    // Add extra arguments if any were specified at config
    const postSlideFunctionArgsRaw =
        documentProperties.getProperty('POST_SLIDE_FUNCTION_ARGS');
    let postSlideFunctionArgs = {};
    if (postSlideFunctionArgsRaw && postSlideFunctionArgsRaw.length > 0) {
      postSlideFunctionArgs = JSON.parse(postSlideFunctionArgsRaw);
    }
    getFunctionByName(postSlideFunction)(slide, row, postSlideFunctionArgs);
  }
}

/**
 * Adds a set of slides by id based on a provided external deck id.
 * It can support both adding slides by ID from a specified deck at
 * configuration level, or adding ALL slides from a variable insight deck.
 * @param {!GoogleAppsScript.Slides.Presentation} deck - The slide deck to add
 *     the slide to.
 * @param {!GoogleAppsScript.Slides.Presentation} insightDeck - The slide deck
 *     to add the slides from.
 * @param {!Array<string>} row Array of strings with information from the
 *     spreadsheet
 */
function addInsightSlides(deck, insightDeck, row) {
  // Add insight slides
  const insightSlidesColumn =
      documentProperties.getProperty('INSIGHT_SLIDE_ID_COLUMN');
  if (insightSlidesColumn && insightSlidesColumn.length > 0) {
    const insights =
        row[insightSlidesColumn - 1].split(',').map((item) => item.trim());
    if (insights.length > 0) {
      let insightDeckToUse;
      let insightSlideIds;
      if (isPresentationId(insights[0])) {
        insightDeckToUse = SlidesApp.openById(insights[0]);
        insightSlideIds =
            insightDeckToUse.getSlides().map((item) => item.getObjectId());
      } else {
        insightDeckToUse = insightDeck;
        insightSlideIds = insights;
      }
      appendInsightSlides(deck, insightDeckToUse, insightSlideIds);
    }
  }
}

/**
 * Determines whether the script should create a slide for this row in the
 * collection based on whether title, subtitle, or body column have been
 * defined. Only one of them should be in order to create a slide.
 *
 * @return {boolean} Whether the script should create a slide for that row in
 *     the collection
 */
function shouldCreateCollectionSlide() {
  const titleColumn = documentProperties.getProperty('TITLE_COLUMN');
  const subtitleColumn = documentProperties.getProperty('SUBTITLE_COLUMN');
  const bodyColumn = documentProperties.getProperty('BODY_COLUMN');
  return (
    (titleColumn && titleColumn.length > 0) ||
      (subtitleColumn && subtitleColumn.length > 0) ||
      (bodyColumn && bodyColumn.length > 0));
}

/**
 * Appends insight slides by reference to the generated deck
 *
 * @param {!Presentation} deck Reference to the generated deck
 * @param {!Presentation} insightDeck Reference to the generated deck
 * @param {!Array<string>} insights Array of slide ids for extended insights
 */
function appendInsightSlides(deck, insightDeck, insights) {
  for (const insightSlideId of insights) {
    if (insightSlideId === '') {
      continue;
    }
    const insightSlide = insightDeck.getSlideById(insightSlideId.trim());
    if (insightSlide === null) {
      continue;
    }
    deck.appendSlide(insightSlide, SlidesApp.SlideLinkingMode.NOT_LINKED);
    if (deck.getMasters().length > 1) {
      deck.getMasters()[deck.getMasters().length - 1].remove();
    }
  }
}

/**
 * Below are the exports required for the linter.
 * This is necessary because AppsScript doesn't support modules.
 */
/* exported retrieveShape */
/* exported appendInsightSlides */
/* exported createDeckFromDatasources */
/* exported replaceSlideShapeWithSheetsChart */
