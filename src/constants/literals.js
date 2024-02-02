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

// Error messages
const ERROR_MISSING_RANGE = 'Couldn\'t find the named range in Configuration.';
const ERROR_NO_SHAPE = 'There was a problem retrieving the shape layout.';
const ERROR_MISSING_VALUE = 'Please select a non-empty cell.';
const ERROR_PARENT_FOLDER =
    'You do not have access to the parent folder of this Sheet.';

// Warning messages
const WARNING_NO_IMAGES = 'No image found for criteria id ';
const WARNING_MULTIPLE_IMAGES = 'No image found for criteria id ';

// Success messages
const SUCCESS_UPLOADED = 'File uploaded for: ';

// Properties configuration
const RANGE_NAME = 'Configuration!PROPERTIES';


/**
 * Below are the exports required for the linter.
 * This is necessary because AppsScript doesn't support modules.
 */
/* exported ERROR_MISSING_RANGE */
/* exported ERROR_NO_SHAPE */
/* exported ERROR_MISSING_VALUE */
/* exported ERROR_PARENT_FOLDER */
/* exported RANGE_NAME */
/* exported WARNING_NO_IMAGES */
/* exported WARNING_MULTIPLE_IMAGES */
/* exported SUCCESS_UPLOADED */

