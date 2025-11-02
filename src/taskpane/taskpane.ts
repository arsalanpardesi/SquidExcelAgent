// src/taskpane/taskpane.ts

// This is required to load the Office.js library
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("runAgent").onclick = runAgentWorkflow;
  }
});

/**
 * Main function that orchestrates the entire AI agent workflow.
 */
async function runAgentWorkflow() {
  const goal = (document.getElementById("goal") as HTMLInputElement).value;
  const sheetHint = (document.getElementById("sheetHint") as HTMLInputElement).value;
  const rowHint = (document.getElementById("rowHint") as HTMLInputElement).value;
  const pdfFileInput = document.getElementById("pdfFile") as HTMLInputElement;
  const statusEl = document.getElementById("status");
  const agentResponseEl = document.getElementById("agentResponse");

  const pdfFile = pdfFileInput.files.length > 0 ? pdfFileInput.files[0] : null;

  if (!goal) {
    agentResponseEl.innerText = "Please enter a command for the agent.";
    return;
  }

  try {
    agentResponseEl.innerHTML = ""; // Clear previous results
    statusEl.innerText = "Reading data and connecting to agent...";

    // This function now gets context from ALL sheets.
    const contextSummary = await getSheetContextForAgent();

    // 1. Prepare FormData for the POST request
    const formData = new FormData();
    formData.append('goal', goal);
    formData.append('context', JSON.stringify(contextSummary));
    formData.append('sheetHint', sheetHint || '');
    formData.append('insertRow', rowHint || '');
    if (pdfFile) {
      formData.append('pdfFile', pdfFile, pdfFile.name);
      formData.append('pdfFileMimeType', pdfFile.type);
    }

    // 2. Use fetch to make the request and get a readable stream
    const response = await fetch("http://localhost:3001/api/agent", {
        method: 'POST',
        body: formData,
    });

    if (!response.ok || !response.body) {
        throw new Error(`Server error: ${response.statusText}`);
    }

    // 3. Process the streaming response
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        
        buffer += decoder.decode(value, { stream: true });
        
        // Parse Server-Sent Events from the buffer
        let eventIndex;
        while ((eventIndex = buffer.indexOf('\n\n')) !== -1) {
            const message = buffer.slice(0, eventIndex);
            buffer = buffer.slice(eventIndex + 2);
            
            let eventType = '';
            const lines = message.split('\n');
            for(const line of lines) {
                if (line.startsWith("event:")) {
                    eventType = line.replace("event:", "").trim();
                } else if (line.startsWith("data:")) {
                    const dataLine = line.replace("data:", "").trim();
                    if (dataLine) {
                        const data = JSON.parse(dataLine);
                        
                        // 4. Handle different event types from the server
                        if (eventType === 'status') {
                            agentResponseEl.innerHTML += `<div>- ${data.replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')}</div>`;
                            statusEl.innerText = data.replace(/\*\*/g, '');
                        }
                        if (eventType === 'summary') {
                            agentResponseEl.innerHTML += `<div style="margin-top: 10px; padding: 8px; background-color: #f0f0f0; border-radius: 4px;"><b>Agent's Plan:</b> ${data}</div>`;
                        }
                        if (eventType === 'subPlan') {
                            statusEl.innerText = "Executing sub-plan in Excel...";
                            await executePlanInExcel({ plans: [data.plan] }); 
                        }
                        if (eventType === 'finalResult') {
                            statusEl.innerText = "Task completed successfully!";
                        }
                         if (eventType === 'error') {
                            throw new Error(data.message);
                        }
                    }
                }
            }
        }
    }

  } catch (error) {
    console.error(error);
    agentResponseEl.innerHTML += `<div style="color: red; font-weight: bold;">An error occurred: ${error.message}</div>`;
    statusEl.innerText = "Finished with errors.";
  }
}

/**
 * Reads data from ALL worksheets and formats it as a summary for the AI agent.
 * This iterates through every sheet in the workbook.
 */
async function getSheetContextForAgent(): Promise<any> {
  let allSheetSummaries = [];

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    for (const sheet of sheets.items) {
      try {
        
        
        const usedRange = sheet.getUsedRange(true);
        usedRange.load("address, rowCount, columnCount, rowIndex, columnIndex");
        
        await context.sync();

        const rowCount = usedRange.rowCount;
        let first3Rows: any[][] = [];
        let last3Rows: any[][] = [];

        if (rowCount > 0) {
          // Case 1: The sheet has more than 6 rows. Get top 3 and bottom 3.
          if (rowCount > 6) {
            const startRow = usedRange.rowIndex;
            const startCol = usedRange.columnIndex;
            const colCount = usedRange.columnCount;

            // Define the range for the first 3 rows
            const first3Range = sheet.getRangeByIndexes(startRow, startCol, 3, colCount);
            first3Range.load("values");

            // Define the range for the last 3 rows
            const last3Range = sheet.getRangeByIndexes(startRow + rowCount - 3, startCol, 3, colCount);
            last3Range.load("values");

            await context.sync();

            first3Rows = first3Range.values;
            last3Rows = last3Range.values;
          }
          // Case 2: The sheet has 6 or fewer rows. Get all of them.
          else {
            usedRange.load("values");
            await context.sync();
            first3Rows = usedRange.values;
            // last3Rows remains empty to avoid sending duplicate data.
          }
        }

        allSheetSummaries.push({
          name: sheet.name,
          usedRangeAddress: usedRange.address,
          totalRows: usedRange.rowCount,
          totalCols: usedRange.columnCount,
          first3Rows: first3Rows,
          last3Rows: last3Rows,
        });
      } catch (error) {
        // If getUsedRange throws (e.g., for a completely empty sheet), catch it.
        console.log(`Worksheet "${sheet.name}" is empty or an error occurred. Adding minimal context.`, error);
        allSheetSummaries.push({
          name: sheet.name,
          usedRangeAddress: "N/A",
          totalRows: 0,
          totalCols: 0,
          first3Rows: [],
          last3Rows: [],
        });
      }
    }
  });

  return { sheets: allSheetSummaries };
}

/**
 * Calls backend server's endpoint to get a plan from the AI. REDUNDANT TO BE DELETED LATER
 */
async function fetchPlanFromAgent(requestData: any): Promise<any> {
  const formData = new FormData();

  formData.append('goal', requestData.goal);
  formData.append('model', requestData.model);
  formData.append('sheetHint', requestData.sheetHint || '');
  formData.append('insertRow', requestData.insertRow || '');
  formData.append('context', JSON.stringify(requestData.context));

  if (requestData.pdfFile) {
    formData.append('pdfFile', requestData.pdfFile, requestData.pdfFile.name);
    
    formData.append('pdfFileMimeType', requestData.pdfFile.type);
  }

  const response = await fetch("http://localhost:3001/api/agent", {
      method: 'POST',
      body: formData
  });

  if (!response.ok) {
      const errorText = await response.text();
      try {
          const errorData = JSON.parse(errorText);
          throw new Error(errorData.error || "Failed to get plan from agent.");
      } catch {
          throw new Error(`Failed to get plan from agent. Server responded with: ${errorText}`);
      }
  }
  
  return response.json();
}


/**
 * Converts a 0-based row and column index to A1-style notation.
 */
function toA1(row: number, col: number): string {
  let colStr = '';
  let n = col + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    colStr = String.fromCharCode(65 + rem) + colStr;
    n = Math.floor((n - 1) / 26);
  }
  return colStr + (row + 1);
}

/**
 * Executes the steps from the AI's generated plans on the workbook.
 * It now handles an array of plans, custom formatting, and global font settings.
 */
export async function executePlanInExcel(result: any) {
  // 1. Check for the new data shape: an object with a 'plans' array.
  if (!result || !result.plans || !Array.isArray(result.plans)) {
    console.log("No valid plans to execute.");
    document.getElementById("status").innerText = "No actionable plan received.";
    return;
  }

  // 2. Loop through each plan returned by the agent.
  for (const plan of result.plans) {
    if (!plan || !Array.isArray(plan)) continue; // Skip empty or invalid sub-plans

    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;

        // 3. Loop through each step within the current plan.
        for (const step of plan) {
          console.log("Executing step:", step);
          document.getElementById("status").innerText = `Executing: ${step.op}...`;

          if (step.op === "createSheet") {
            const sheet = workbook.worksheets.add(step.args.name);
            sheet.load('name');
          } else {
            const rangeRef = step.args.range;
            if (!rangeRef || !rangeRef.sheet) {
                console.warn("Skipping step due to missing range or sheet:", step);
                continue;
            }

            const sheet = workbook.worksheets.getItem(rangeRef.sheet);
            const startCell = toA1(rangeRef.r1, rangeRef.c1);
            const endCell = toA1(rangeRef.r2, rangeRef.c2);
            const address = startCell === endCell ? startCell : `${startCell}:${endCell}`;
            const range = sheet.getRange(address);

            // --- NEW: Global Font Formatting ---
            range.format.font.name = "Arial";
            range.format.font.size = 10;
            // --- END OF NEW FONT FORMATTING ---

            switch (step.op) {
              case "setValues":
                range.values = step.args.values;

                // Header row formatting - TO BE CONSIDERED LATER
                //if (rangeRef.r1 === 0) {
                //  range.format.fill.color = "#4472C4";
                //  range.format.font.color = "white";
                //  range.format.font.bold = true;
                //}

                // Total/Subtotal row formatting - TO BE CONSIDERED LATER
                //const firstCellValue = step.args.values[0]?.[0];
                //if (typeof firstCellValue === 'string') {
                //  const lowerCaseValue = firstCellValue.toLowerCase();
                //  const isTotalRow = ['total', 'subtotal', 'net income', 'gross profit'].some(keyword => lowerCaseValue.includes(keyword));
                  
                //  if (isTotalRow) {
                //    range.format.borders.getItem('EdgeTop').style = 'Continuous';
                //    range.format.borders.getItem('EdgeTop').weight = 'Thin';
                //    range.format.fill.color = "#4472C4";
                //    range.format.font.color = "white";
                //    range.format.font.bold = true;
                //    range.format.borders.getItem('EdgeBottom').style = 'Double';
                //    range.format.borders.getItem('EdgeBottom').weight = 'Thick';
                //  }
                //}
                break;

              case "setFormulas":
                range.formulas = step.args.formulas;
                break;
              
              case "formatRange":
                if (step.args.format === 'percent') range.numberFormat = [["0.00%"]];
                if (step.args.format === 'currency') range.numberFormat = [["$#,##0.00"]];
                // --- NEW: Zero Decimal Formatting ---
                if (step.args.format === 'zero_decimal') range.numberFormat = [["0"]];
                // --- END OF NEW FORMATTING ---
                break;
            }
          }
        }

        await context.sync();
      });
      console.log("Successfully executed a plan batch.");

    } catch (error) {
      console.error("Failed to execute a plan batch. Error:", error);
      document.getElementById("status").innerText = `Error during execution. Check console.`;
      throw new Error(`Execution failed: ${error.message}`);
    }
  }
}