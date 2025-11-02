// This file will run before all tests, ensuring the environment is set up.
process.env.GEMINI_API_KEY = 'mock-api-key-for-testing';

// This file will run before all tests, ensuring the global Office environment is mocked.

// --- Create Mocks for the Excel Object Model ---

const mockRange = {
  load: jest.fn(),
  values: [],
  formulas: [],
  numberFormat: [],
  format: {
    font: { name: '', size: 0, color: '', bold: false },
    fill: { color: '' },
    borders: { 
      getItem: jest.fn().mockReturnThis(), // Returns the border object
      style: '',
      weight: '',
    },
  },
};

const mockSheet = {
  load: jest.fn(),
  getRange: jest.fn(() => mockRange), // getRange returns our mockRange
  name: 'MockSheet',
};

const mockContext = {
  workbook: {
    worksheets: {
      add: jest.fn(() => mockSheet), // add returns our mockSheet
      getItem: jest.fn(() => mockSheet), // getItem returns our mockSheet
    },
  },
  sync: jest.fn(),
};

// --- Assign Mocks to the Global Scope ---

global.Office = {
  onReady: jest.fn((callback) => callback({ host: 'Excel' })),
  HostType: { Excel: 'Excel' },
} as any;

global.Excel = {
  run: jest.fn(async (callback) => {
    // When Excel.run is called, it executes the callback with our mock context.
    await callback(mockContext);
  }),
} as any;

global.document = {
  getElementById: jest.fn().mockReturnValue({ 
    innerText: '',
    innerHTML: '',
    value: '',
    files: [],
  }),
} as any;

// Export the mocks so we can inspect them in our tests
export { mockContext, mockSheet, mockRange };