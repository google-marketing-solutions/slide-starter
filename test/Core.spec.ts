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

import { assert } from "chai";
import "mocha";
import rewire from "rewire";

globalThis.PropertiesService = {
  getDocumentProperties: () => {
    return {
      getProperty: (id: string) => {
        if (id === "LAYOUT_NAME") {
          return "test-name";
        }
      },
    };
  },
};

const core = rewire("../src/sls_facades.js");
const getTemplateLayoutId = core.__get__("getTemplateLayoutId");

describe("Core library", function () {
  it("should getTemplateLayoutId", function () {
    globalThis.Slides = {
      Presentations: {
        get: (_: string) => {
          return {
            layouts: [
              {
                objectId: "test-id",
                layoutProperties: {
                  displayName: "test-name",
                },
              },
            ],
          };
        },
      },
    };
    const id = getTemplateLayoutId("test-id");
    assert.equal(id, "test-id");
  });
});
