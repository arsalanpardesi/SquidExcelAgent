import { SheetModel } from './sheet.js';

describe('SheetModel', () => {

  it('should create a sheet, set values and formulas, and evaluate correctly', () => {
    const model = new SheetModel();
    // The model starts with 'Sheet1', let's delete it and start fresh for a clean test
    model.deleteSheet('Sheet1');
    model.createSheet('TestSheet');
    
    // Set some initial values
    model.setValues({ sheet: 'TestSheet', r1: 0, c1: 0, r2: 0, c2: 1 }, [['Revenue', 1000]]);
    model.setValues({ sheet: 'TestSheet', r1: 1, c1: 0, r2: 1, c2: 1 }, [['COGS', 600]]);

    // Set a formula that depends on the values
    model.setFormulas({ sheet: 'TestSheet', r1: 2, c1: 0, r2: 2, c2: 1 }, [['Gross Profit', '=B1-B2']]);
    
    // Evaluate all formulas
    model.evaluateAll();

    const result = model.toJSON();
    const testSheet = result.sheets.find(s => s.name === 'TestSheet');
    
    // Assert that the formula was calculated correctly
    expect(testSheet).toBeDefined();
    expect(testSheet!.rows[2][1].value).toBe(400); // 1000 - 600 = 400
  });

  it('should correctly load a workbook with formulas from JSON', () => {
    // This tests the bug we fixed!
    const workbookData = {
      sheets: [
        {
          name: 'MySheet',
          rows: [
            [{ value: 10 }, { value: 20 }],
            [{ value: null, formula: '=A1+B1' }],
          ],
        },
      ],
    };
    
    const model = SheetModel.fromJSON(workbookData);

    const result = model.toJSON();
    const sheet = result.sheets[0];

    // Assert that the formula was loaded AND evaluated
    expect(sheet.rows[1][0].value).toBe(30);
  });
  describe('Error Handling', () => {
    it('should throw an error when creating a duplicate sheet', () => {
      const model = new SheetModel();
      model.createSheet('DuplicateSheet');
      // Expecting the second call with the same name to throw an error
      expect(() => model.createSheet('DuplicateSheet')).toThrow('Sheet exists');
    });

    it('should throw an error when getting a non-existent sheet', () => {
      const model = new SheetModel();
      expect(() => model.dispatch('setValues', { range: { sheet: 'GhostSheet', r1: 0, c1: 0, r2: 0, c2: 0 } }))
        .toThrow('Sheet not found: GhostSheet');
    });
  });

  describe('Undo Functionality', () => {
    it('should revert the last operation when undo is called', () => {
      const model = new SheetModel();
      model.setValues({ sheet: 'Sheet1', r1: 0, c1: 0, r2: 0, c2: 0 }, [['Initial Value']]);
      
      // Perform an action
      model.setValues({ sheet: 'Sheet1', r1: 0, c1: 0, r2: 0, c2: 0 }, [['New Value']]);
      let sheetState = model.toJSON().sheets[0];
      expect(sheetState.rows[0][0].value).toBe('New Value');

      // Perform undo
      model.undo();
      sheetState = model.toJSON().sheets[0];
      // Assert that the state has reverted
      expect(sheetState.rows[0][0].value).toBe('Initial Value');
    });
  });

  describe('Formula Evaluation Edge Cases', () => {
    it('should return #REF! for a circular reference', () => {
      const model = new SheetModel();
      model.setFormulas({ sheet: 'Sheet1', r1: 0, c1: 0, r2: 0, c2: 0 }, [['=A1']]);
      model.evaluateAll();
      const result = model.toJSON();
      expect(result.sheets[0].rows[0][0].value).toBe('#REF!');
    });

    it('should return #ERROR! for a malformed formula', () => {
      const model = new SheetModel();
      model.setFormulas({ sheet: 'Sheet1', r1: 0, c1: 0, r2: 0, c2: 0 }, [['=1+/2']]);
      model.evaluateAll();
      const result = model.toJSON();
      expect(result.sheets[0].rows[0][0].value).toBe('#ERROR!');
    });
  });
  
  describe('A1 Notation Helpers', () => {
    it('should convert cell coordinates to A1 notation', () => {
      expect(SheetModel.rcToA1(0, 0)).toBe('A1');
      expect(SheetModel.rcToA1(9, 25)).toBe('Z10');
      expect(SheetModel.rcToA1(0, 26)).toBe('AA1');
    });

    it('should convert A1 notation to cell coordinates', () => {
      expect(SheetModel.a1ToRc('A1')).toEqual({ r: 0, c: 0 });
      expect(SheetModel.a1ToRc('Z10')).toEqual({ r: 9, c: 25 });
      expect(SheetModel.a1ToRc('AA1')).toEqual({ r: 0, c: 26 });
    });
  });
  
  it('should propagate #REF! errors to dependent cells', () => {
    const model = new SheetModel();
    model.deleteSheet('Sheet1');
    model.createSheet('ErrorSheet');

    // Create a circular reference in A1
    model.setFormulas({ sheet: 'ErrorSheet', r1: 0, c1: 0, r2: 0, c2: 0 }, [['=A1']]);
    // Create a formula in B1 that depends on A1
    model.setFormulas({ sheet: 'ErrorSheet', r1: 0, c1: 1, r2: 0, c2: 1 }, [['=A1+10']]);

    model.evaluateAll();
    const result = model.toJSON();
    const sheet = result.sheets[0];

    // Assert that the original error is correct
    expect(sheet.rows[0][0].value).toBe('#REF!');
    // Assert that the error has propagated to the dependent cell
    expect(sheet.rows[0][1].value).toBe('#REF!');
  });
});