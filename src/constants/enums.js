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
 * Determines a color based on if a value is a Good, Needs Improvement or Poor
 * range for a given metric.
 *
 * @param {!Array} range Array with a low and high threshold for a CWV metric
 * @param {number} value Number indicating the metric score
 * @return {!Array} Array of RBG values in decimal form
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
 * Object whose keys represent Core Web Vital metrics and values are Arrays that
 * contain the low & high thresholds for that metric. Used in coloring the table
 * for CrUX CWV data.
 */
const CWV = {
  LCP: [2500, 4000],
  FID: [100, 300],
  INP: [100, 300],
  CLS: [0.1, 0.25],
};

const cwvTextType = {
  'CRUX_FID': CWV.FID,
  'CRUX_CLS': CWV.CLS,
  'CRUX_INP': CWV.INP,
  'CRUX_LCP': CWV.LCP,
};

/**
 * Object whose keys are colors and values are arrays of RGB values in decimal.
 * Used in coloring the table for CrUX CWV data.
 */
const COLORS = {
  GREEN: '#34A853', // Good
  YELLOW: '#FBBC04', // Needs Improvement
  RED: '#EA4335', // Poor
  WHITE: '#F8F9FA', // None
};
