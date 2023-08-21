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
