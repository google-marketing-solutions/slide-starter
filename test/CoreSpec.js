globalThis.PropertiesService = {
    getDocumentProperties: () => {
        return {
            getProperty: (id) => {
                if (id === 'LAYOUT_NAME') {
                    return 'test-name'
                }
            }
        }
    }
};

const core = require("../src/sls_core.js")

describe("Sls Core", function () {
    it("should getTemplateLayoutId", function () {
        globalThis.Slides = {
            Presentations: {
                get: (presentationId) => {
                    return {
                        layouts: [
                            {
                                objectId: 'test-id',
                                layoutProperties: {
                                    displayName: 'test-name'
                                }
                            }
                        ]
                    }
                }
            }
        };
        const id = core.getTemplateLayoutId('test-id');
        expect(id).toEqual('test-id');
    });
});
