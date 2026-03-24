/**
 * Jest setup: stub the Office JS global so lib modules can be imported
 * in a Node/jsdom environment without a running PowerPoint host.
 *
 * office-addin-mock provides a partial mock; we supplement with minimal stubs
 * needed for our specific API surface.
 */

// Provide a minimal Office global so imports don't throw at module load time.
// Full mock objects are set up per-test using office-addin-mock.

const officeMock = {
  AsyncResultStatus: {
    Succeeded: "succeeded",
    Failed: "failed",
  },
  EventType: {
    DocumentSelectionChanged: "documentSelectionChanged",
  },
  context: {
    document: {
      customXmlParts: {
        getByNamespaceAsync: jest.fn(),
        addAsync: jest.fn(),
      },
      addHandlerAsync: jest.fn(),
      removeHandlerAsync: jest.fn(),
    },
  },
  onReady: (cb: () => void) => cb(),
};

// @ts-expect-error — global stub for test environment
global.Office = officeMock;

// Minimal PowerPoint global stub (extended per-test as needed)
// @ts-expect-error
global.PowerPoint = {
  run: jest.fn(),
};
