import { executePlanInExcel } from './taskpane.js';
// Import the mocks from the setup file to inspect them
import { mockContext, mockSheet, mockRange } from '../../jest.setup.js';

describe('executePlanInExcel', () => {

  beforeEach(() => {
    // Reset the call history of all mock functions before each test
    jest.clearAllMocks();
  });

  it('should execute a multi-step, multi-plan workflow sequentially', async () => {
    // Arrange
    const complexResult = {
      plans: [
        [ { op: 'createSheet', args: { name: 'Financials' } } ],
        [ { op: 'setFormulas', args: { range: { sheet: 'Financials', r1: 1, c1: 0, r2: 1, c2: 0 }, formulas: [['=SUM(A1)']] } } ]
      ]
    };

    // Act
    await executePlanInExcel(complexResult);

    // Assert
    expect(Excel.run).toHaveBeenCalledTimes(2);
    // Check that the createSheet operation was called correctly
    expect(mockContext.workbook.worksheets.add).toHaveBeenCalledWith('Financials');
    // Check that the setFormulas operation was called correctly
    expect(mockSheet.getRange).toHaveBeenCalledWith('A2');
    expect(mockRange.formulas).toEqual([['=SUM(A1)']]);
  });

  it('should apply all custom formatting rules correctly', async () => {
    // Arrange
    const formattingResult = {
      plans: [[
        { op: 'setValues', args: { range: { sheet: 'MySheet', r1: 0, c1: 0, r2: 0, c2: 0 }, values: [['Header']] } },
        { op: 'formatRange', args: { range: { sheet: 'MySheet', r1: 2, c1: 0, r2: 2, c2: 0 }, format: 'currency' } }
      ]]
    };

    // Act
    await executePlanInExcel(formattingResult);

    // Assert
    expect(Excel.run).toHaveBeenCalledTimes(1);
    expect(mockSheet.getRange).toHaveBeenCalledWith('A1');
    expect(mockSheet.getRange).toHaveBeenCalledWith('A3');
    expect(mockRange.format.font.name).toBe('Arial');
    expect(mockRange.format.font.size).toBe(10);
    expect(mockRange.numberFormat).toEqual([["$#,##0.00"]]);
  });
});