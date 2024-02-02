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
 * Checks if the provided image value is a base64 encoded image.
 * @param {string} imageValue - The string value representing the image.
 * @return {boolean} True if the image is base64 encoded, False otherwise.
 */
function isBase64Image(imageValue) {
  return imageValue.match(/^data:image\/([a-z]+);base64,/i);
}

/**
 * Decodes a base64 encoded image string and returns a blob.
 * @param {string} imageValue - The base64 encoded image string.
 * @return {Blob} A blob object containing the decoded image data.
 */
function decodeBase64Image(imageValue) {
  const match = imageValue.match(/^data:image\/([a-z]+);base64,/i);
  const imageType = match[1];
  const imageBase64 = imageValue.split(',')[1];
  const decodedImage = Utilities.base64Decode(imageBase64);
  return Utilities.newBlob(decodedImage, MimeType[imageType.toUpperCase()]);
}

/**
 * Gets a function by name.
 *
 * @param {string} functionName The name of the function to get.
 * @return {!Function} The function with the given name.
 * @throws {Error} If the function name is not alphanumeric.
 */
function getFunctionByName(functionName) {
  const alphanumericRegex = /^[a-zA-Z0-9]+$/;
  if (!alphanumericRegex.test(functionName)) {
    throw new Error('Function name not alphanumeric');
  }
  return new Function(`return ${functionName};`)();
}

/**
 * Checks if the given URL is a valid image URL.
 *
 * @param {string} url The URL to check.
 * @return {boolean} Whether the URL is a valid image URL.
 */
function isValidImageUrl(url) {
  // TODO: Fix issue #39 Check if it's an image
  return url.startsWith('http://') || url.startsWith('https://');
}