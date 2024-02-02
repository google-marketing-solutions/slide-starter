/* eslint-disable */
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

import { assert, expect } from "chai";
import "mocha";
import rewire from "rewire";

const helpers = rewire('../src/utilities/helpers.js');
const isValidImageUrl = helpers.__get__('isValidImageUrl');
const getFunctionByName = helpers.__get__('getFunctionByName');

describe('isValidImageUrl', () => {
    it('should return true for a valid image URL', () => {
      const url = 'https://example.com/image.png';
      const isValid = isValidImageUrl(url);
      assert.isTrue(isValid);
    });
  
    it('should return false for an invalid image URL', () => {
      const url = 'mailto:example@example.com';
      const isValid = isValidImageUrl(url);
      assert.isFalse(isValid);
    });
  });

  describe('getFunctionByName', () => {
    it('should return the function with the given name', () => {
      const functionName = 'myFunction';
      const functionValue = () => {};
      globalThis[functionName] = functionValue;
  
      const functionReturned = getFunctionByName(functionName);
  
      assert.strictEqual(functionReturned, functionValue);
    });
  
    it('should throw an error if the function name is not alphanumeric', () => {
      const functionName = 'myFunction123!';
  
      expect(() => {
        getFunctionByName(functionName);
      }).to.throw(Error, 'Function name not alphanumeric');
    });
  });