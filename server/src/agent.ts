import dotenv from 'dotenv';
import { Plan, PlanStep } from './types.js';
import { SheetModel } from './sheet.js';
import { GoogleGenAI } from '@google/genai';
import { StateGraph, END, Annotation, START} from '@langchain/langgraph';
import { z } from 'zod';
import { MemorySaver } from '@langchain/langgraph';

dotenv.config();

// --- Ollama Configuration ---
const OLLAMA_BASE_URL = process.env.OLLAMA_BASE_URL || 'http://localhost:11434';
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'qwen3:32b';

// --- Gemini Configuration & Initialization ---
export const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const geminiAI = GEMINI_API_KEY ? new GoogleGenAI({apiKey: GEMINI_API_KEY}) : null;

export type AgentHints = { sheetHint?: string; insertRow?: number }; // insertRow is 0-based

// --- Gemini File Upload Function ---
// This function takes the PDF buffer, uploads it, and returns the file part for the prompt.
async function uploadPdfToGemini(pdfFile: { buffer: Buffer; mimetype: string }) {
  if (!geminiAI) {
    throw new Error('Gemini API key not configured.');
  }
  
  console.log(`Uploading ${pdfFile.mimetype} to Gemini File API...`);
  
  const uint8Array = new Uint8Array(pdfFile.buffer);
  const blob = new Blob([uint8Array], { type: pdfFile.mimetype });

  const result = await geminiAI.files.upload({
      file: blob,
      config: {
          mimeType: pdfFile.mimetype,
      }
  });
  
  console.log(`File uploaded successfully. URI: ${result.uri}`);
  
  return {
    fileData: {
      mimeType: result.mimeType,
      fileUri: result.uri,
    },
  };
}

/* ---------- Context summarizer sent to the model ---------- */
function getCellText(cell: any) {
  if (!cell) return '';
  return cell.formula ? cell.formula : (cell.value ?? '');
}

function summarizeWorkbook(wb: SheetModel) {
  const out: any = { sheets: [] };

  for (const s of wb.wb.sheets.values()) {
    const rows = s.rows ?? [];
    const rowCount = rows.length;
    const colCount = rows.reduce((m: number, r: any[]) => Math.max(m, r?.length ?? 0), 0);

    const prevRows = Math.min(rowCount, 60);
    const prevCols = Math.min(colCount, 20);
    const previewA1: any[] = [];
    for (let r = 0; r < prevRows; r++) {
      const row: any[] = [];
      for (let c = 0; c < prevCols; c++) row.push(getCellText(rows[r]?.[c]));
      previewA1.push(row);
    }

    const headerA1: any[] = [];
    if (rowCount > 0) {
      for (let c = 0; c < Math.min(colCount, 50); c++) headerA1.push(getCellText(rows[0]?.[c]));
    }

    const labelsA: { row: number; text: string }[] = [];
    const capRows = Math.min(rowCount, 500);
    for (let r = 0; r < capRows; r++) labelsA.push({ row: r, text: String(getCellText(rows[r]?.[0] ?? '')).trim() });

    out.sheets.push({
      name: s.name,
      rows: rowCount,
      cols: colCount,
      headerA1,
      labelsA,
      previewA1
    });
  }
  return out;
}

/* ---------- System prompt / tool contract ---------- */
const SYSTEM_PROMPT = `
You are an AI spreadsheet operator. Your task is to generate a JSON plan to fulfill a user's goal.

You receive:
- a user's goal,
- optional hints: { "sheetHint"?: string, "insertRow"?: number },
- and a JSON summary of the workbook, including for each sheet:
  - headerA1: first row cells (helps identify period columns),
  - labelsA: [{row, text}] for up to the first 500 rows of column A (captions),
  - previewA1: a small grid preview.

Return ONLY a single, root-level JSON object and nothing else (no prose, no markdown formatting). The object MUST be:
{ "plan": PlanStep[], "summary": string }

PlanStep is exactly one of:
- { "op": "createSheet",   "args": { "name": string } }
- { "op": "setValues",     "args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "values": (string|number|null)[][] } }
- { "op": "setFormulas",   "args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "formulas": (string|null)[][] } }
- { "op": "formatRange",   "args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "format": "percent" | "currency" | "number" | "text" } }

Rules:
- **Sheet Creation**: You MUST use the 'createSheet' operation for any new sheet BEFORE you can use 'setValues' or other operations on that sheet. Check if a sheet exists in the context before creating a new one.
- **Indexing is 0-based**: All row/column numbers you receive (in labelsA) and provide (in ranges) are 0-based. HOWEVER, when writing Excel-style formulas (e.g., "=A1+B1"), you MUST convert the 0-based row index to a 1-based row number by adding 1. For example, to reference a cell at row index 4, use "5" in the formula string (e.g., "A5").
- Ensure 2D array sizes for values/formulas exactly match the range dimensions.
- If "sheetHint" is provided, prefer that exact sheet when writing unless it clearly doesn't exist.
- If "insertRow" is provided, place new calculation rows starting at that row index (0-based). Expand the sheet if necessary.
- **Captions**: For any new calculated row, ALWAYS write a descriptive caption in column A (c1==0), e.g., "Gross margin %".
- **Use captions to find inputs**: Use the sheet's labelsA (case-insensitive) to find row indices for:
    - Revenue: ["revenue", "net sales", "sales"]
    - Cost of revenue: ["cost of revenue", "cost of sales", "cogs"]
    - Gross profit: ["gross profit"]
  Prefer exact/starts-with matches, ignoring punctuation/whitespace.
- **Use headerA1** to align period columns (usually headers start at column 1). When writing across years, place formulas in the same data columns as Revenue etc.
- Write formulas, not hard-coded numbers, e.g.:
    - Gross margin % = Gross profit / Revenue
    - Net margin %   = Net income / Revenue
- When calculating percentage metrics, set "formatRange" to "percent" for the data columns you write (do not format the label cell).
- **Keep plans under 50 steps** to ensure efficiency and accuracy.
`.trim();

/* ---------- Utilities ---------- */
function messagesToPrompt(messages: { role: string; content: string }[]) {
  return messages.map(m => `${m.role.toUpperCase()}: ${m.content}`).join('\n\n');
}

function cleanToJsonString(raw: string) {
  return String(raw)
    .replace(/^\uFEFF/, '')
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/```$/i, '')
    .trim();
}

/* ---------- Streaming (SSE) ---------- */
export type AgentStreamEvent =
  | { type: 'status'; data: string }
  | { type: 'context'; data: any }
  | { type: 'token'; data: string }
  | { type: 'plan'; data: any }
  | { type: 'error'; data: string }
  | { type: 'done'; data: string };

// --- Gemini Streaming Function ---
async function streamFromGemini(messages: any[], on: (e: AgentStreamEvent) => void): Promise<string> {
    if (!geminiAI) {
      throw new Error('Gemini API key not configured. Please check your .env file.');
    }
  
    const userPrompt = messages.find(m => m.role === 'user')?.content;
    if (!userPrompt) {
        throw new Error('No user content found in messages for Gemini.');
    }
  
    const result = await geminiAI.models.generateContentStream({
      model: "gemini-2.5-pro", // Specify the model here
      contents: [{ role: "user", parts: [{ text: userPrompt }] }],
      config: {
        systemInstruction: [{ text: SYSTEM_PROMPT }],
        responseMimeType: "application/json",
        temperature: 0,
      }
    });
  
    let fullResponse = '';
    for await (const chunk of result) {
      const chunkText = chunk.text;
      if (chunkText) {
        on({ type: 'token', data: chunkText });
        fullResponse += chunkText;
      }
    }
    return fullResponse;
}

// --- Ollama Streaming Function ---
async function streamFromOllama(messages: any[], on: (e: AgentStreamEvent) => void): Promise<string> {
  const bodyChat: any = {
    model: OLLAMA_MODEL,
    messages,
    stream: true,
    format: 'json',
    options: { temperature: 0 }
  };

  let res: Response;
  try {
    res = await fetch(`${OLLAMA_BASE_URL}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyChat)
    });
  } catch (e) {
    throw new Error(`Could not reach Ollama at ${OLLAMA_BASE_URL} – ${String(e)}`);
  }

  if (res.status === 404) {
    const bodyGen: any = {
      model: OLLAMA_MODEL,
      prompt: messagesToPrompt(messages),
      stream: true,
      format: 'json',
      options: { temperature: 0 }
    };
    const r2 = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyGen)
    });
    if (!r2.ok) throw new Error(`Ollama /api/generate ${r2.status}`);
    return await streamRead(r2, on, 'generate');
  }

  if (!res.ok) throw new Error(`Ollama /api/chat ${res.status}`);
  return await streamRead(res, on, 'chat');
}

async function streamRead(res: Response, on: (e: AgentStreamEvent)=>void, mode: 'chat'|'generate'): Promise<string> {
  const reader = res.body!.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  let full = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    const lines = buf.split('\n');
    buf = lines.pop() || '';
    for (const line of lines) {
      const s = line.trim();
      if (!s) continue;
      try {
        const j = JSON.parse(s);
        const token = mode === 'chat' ? (j?.message?.content ?? '') : (j?.response ?? '');
        if (token) { on({ type: 'token', data: token }); full += token; }
        if (j?.done) break;
      } catch {
        // ignore partials
      }
    }
  }
  return full;
}

/* ---------- Public APIs ---------- */
export async function streamPlanAndExecute(
  goal: string,
  wb: SheetModel,
  hints: AgentHints | undefined,
  on: (e: AgentStreamEvent) => void,
  modelChoice: 'ollama' | 'gemini' // Parameter to select the model
) {
  const context = summarizeWorkbook(wb);
  on({ type: 'context', data: context });

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'user', content: JSON.stringify({ goal, hints: hints ?? {}, context }) }
  ];

  on({ type: 'status', data: `contacting ${modelChoice} model...` });
  
  let raw: string;
  // Logic to switch between models
  if (modelChoice === 'gemini') {
    raw = await streamFromGemini(messages, on);
  } else {
    raw = await streamFromOllama(messages, on);
  }

  on({ type: 'status', data: 'parsing plan...' });

  let plan: Plan;
  try {
    plan = JSON.parse(cleanToJsonString(raw));
  } catch (e) {
    on({ type: 'error', data: `Model returned invalid JSON. Raw output: ${raw}` });
    throw e;
  }
  on({ type: 'plan', data: plan });

  const executed = executePlan(plan, wb);
  on({ type: 'status', data: `executed ${executed.length} step(s)` });
  return { plan: { ...plan, steps: executed }, workbook: wb.toJSON() };
}

function executePlan(plan: Plan, wb: SheetModel) {
  const executed: any[] = [];
  for (const step of plan.steps || []) {
    try {
      const sanitizedArgs = JSON.parse(JSON.stringify(step.args));
      if (sanitizedArgs.range) {
        if (sanitizedArgs.range.r3 !== undefined && sanitizedArgs.range.r2 === undefined) {
          sanitizedArgs.range.r2 = sanitizedArgs.range.r3;
          delete sanitizedArgs.range.r3;
        }
        if (sanitizedArgs.range.c3 !== undefined && sanitizedArgs.range.c2 === undefined) {
          sanitizedArgs.range.c2 = sanitizedArgs.range.c3;
          delete sanitizedArgs.range.c3;
        }
      }

      wb.dispatch(step.op, sanitizedArgs);
      executed.push(step);
    } catch (err) {
      executed.push({ ...step, explain: `FAILED: ${(err as Error).message}` });
      break;
    }
  }
  wb.evaluateAll();
  wb.checkpoint('agent');
  return executed;
}

/**
 * A non-streaming function specifically for the Excel add-in.
 * It accepts a context summary directly from the add-in and uses the established API patterns.
 */
export async function planAndExecuteForAddin(
  goal: string,
  hints: AgentHints | undefined,
  context: any, // The summary object sent from the add-in
  modelChoice: 'ollama' | 'gemini',
  pdfFile?: { buffer: Buffer; mimetype: string } // The file object from the server
) {
  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'user', content: JSON.stringify({ goal, hints: hints ?? {}, context }) }
  ];

  let raw: string;

  if (modelChoice === 'gemini') {
    if (!geminiAI) {
      throw new Error("Gemini API key not configured. Please check your .env file.");
    }
    const userPrompt = messages.find(m => m.role === 'user')?.content;
    if (!userPrompt) {
        throw new Error('No user content found in messages for Gemini.');
    }

    // Prepare the parts for the Gemini prompt.
    const promptParts: any[] = [ { text: userPrompt } ];

    // If a PDF file is provided, upload it using the corrected function
    // and add the resulting file part to the prompt.
    if (pdfFile) {
      const filePart = await uploadPdfToGemini(pdfFile);
      promptParts.unshift(filePart); // Add file at the beginning
    }

    // Using the non-streaming equivalent of your working API call
    const response_ns = await geminiAI.models.generateContent({
      model: "gemini-2.5-pro", // Using the latest pro model
      contents: [{ role: "user", parts: promptParts }],
      // Replicating exact 'config' object structure
      config: {
        systemInstruction: [{ text: SYSTEM_PROMPT }],
        responseMimeType: "application/json",
        temperature: 0,
      }
    });
    

    raw = response_ns.text ?? '{}';

  } else {
    //Note to self: Ollama pdf ability to be added later
    if (pdfFile) {
      console.log("PDF file was provided but will be ignored for the Ollama model.");
    }
    // Standard non-streaming call for Ollama
    const body = { model: OLLAMA_MODEL, messages, stream: false, format: 'json' };
    const res = await fetch(`${OLLAMA_BASE_URL}/api/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });
    if (!res.ok) {
        const errorText = await res.text();
        throw new Error(`Ollama request failed: ${errorText}`);
    }
    const j = await res.json();
    raw = j?.message?.content ?? '{}';
  }

  const plan = JSON.parse(cleanToJsonString(raw));
  
  return { plan };
}


// Agentic workflow ----->

// =================================================================
// 1. AGENT STATE DEFINITION
// =================================================================

const PlanStepSchema = z.object({
  op: z.string(),
  args: z.record(z.any()),
  explain: z.string().optional(),
});

type PlanStepType = z.infer<typeof PlanStepSchema>;

const AgentStateAnnotation = Annotation.Root({
  userInput: Annotation<string>(),
  workbookContext: Annotation<any>(),
  hints: Annotation<AgentHints>(),
  routingDecision: Annotation<'simple' | 'complex' | null>(),
  subTasks: Annotation<string[]>(),
  currentSubTaskIndex: Annotation<number>(),
  finalPlans: Annotation<PlanStepType[][]>(), 
  planHistory: Annotation<PlanStepType[][]>(),
  planSummary: Annotation<string | null>(),
  pdfFile: Annotation<any>(),
  statusUpdates: Annotation<string[]>(),
});

export type AgentState = typeof AgentStateAnnotation.State;

// =================================================================
// 2. GRAPH NODE IMPLEMENTATIONS
// =================================================================

// ### Router Node ###
// Decides if the task is simple (one plan) or complex (multiple sub-plans).
const routerNode = async (state: AgentState): Promise<Partial<AgentState>> => {
    console.log("--- Executing Router Node ---");
    if (!geminiAI) throw new Error("Gemini API key not configured.");

    const prompt = `
        You are a triage agent. Your job is to determine if a user's request for a spreadsheet task is "simple" or "complex".
        - A "simple" task can be fully planned in a single, straightforward series of steps (e.g., a basic mortgage calculation, summing a column).
        - A "complex" task requires breaking the problem down into distinct logical parts that must be planned sequentially (e.g., "Analyze this 10-K, build a financial model, and then create a summary dashboard").

        Based on the user's goal below, respond with a single JSON object with one key, "decision", set to either "simple" or "complex".

        User Goal: "${state.userInput}"
    `;

    const response = await geminiAI.models.generateContent({
        model: "gemini-2.0-flash",
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        config: { responseMimeType: "application/json", temperature: 0 }
    });
    
    const result = JSON.parse(response.text ?? '{}');
    console.log(`Routing decision: ${result.decision}`);
    return { 
        routingDecision: result.decision,
        statusUpdates: [`Task complexity assessed as: **${result.decision}**.`]
    };
};


// ### Planner for SIMPLE tasks ###
const singlePlannerNode = async (state: AgentState): Promise<Partial<AgentState>> => {
  console.log("--- Executing Single Planner Node (Simple Task) ---");
  if (!geminiAI) throw new Error("Gemini API key not configured.");

  const promptParts: any[] = [{
      text: `Based on the user's goal, the provided spreadsheet context, and the content of the attached file (if any), create a complete, detailed, step-by-step JSON plan. User Goal: ${state.userInput}\nSpreadsheet Context: ${JSON.stringify(state.workbookContext)}`
  }];
  if (state.pdfFile) {
      const filePart = await uploadPdfToGemini(state.pdfFile);
      promptParts.unshift(filePart);
  }
  const response = await geminiAI.models.generateContent({
      model: "gemini-2.5-pro",
      contents: [{ role: "user", parts: promptParts }],
      config: {
        systemInstruction: [{ text: `Your output MUST be a single JSON object with "plan" (an array of PlanStep objects) and "summary" keys. Schema: ${SYSTEM_PROMPT}` }],
        responseMimeType: "application/json", temperature: 0
      }
  });

  const result = JSON.parse(response.text ?? '{}');
  return { 
      finalPlans: [result.plan || []],
      planSummary: result.summary || null, 
      statusUpdates: [`Comprehensive plan created with ${result.plan?.length || 0} steps.`]
  };
};

// ### Nodes for COMPLEX tasks ###
const taskAllocatorNode = async (state: AgentState): Promise<Partial<AgentState>> => {
    console.log("--- Executing Task Allocator Node (Complex Task) ---");
    if (!geminiAI) throw new Error("Gemini API key not configured.");
    
    const promptParts: any[] = [{
        text: `You are an expert project manager. Based on the user's goal and the content of the attached file (if any), break down the complex user request into a series of 3-20 sequential sub-tasks. Return a JSON object with a single key "subTasks", an array of strings. User Goal: "${state.userInput}".`
    }];
    if (state.pdfFile) {
        const filePart = await uploadPdfToGemini(state.pdfFile);
        promptParts.unshift(filePart);
    }    
    
    const response = await geminiAI.models.generateContent({
        model: "gemini-2.5-pro",
        contents: [{ role: "user", parts: promptParts }],
        config: { responseMimeType: "application/json", temperature: 0 }
    });

    const result = JSON.parse(response.text ?? '{}');
    return { 
        subTasks: result.subTasks || [], 
        currentSubTaskIndex: 0,
        finalPlans: [], // Initialize the list of plans
        statusUpdates: [`Task broken down into ${result.subTasks?.length || 0} sub-tasks.`]
    };
};

const subPlannerNode = async (state: AgentState): Promise<Partial<AgentState>> => {
    const currentTask = state.subTasks[state.currentSubTaskIndex];
    console.log(`--- Planning Sub-Task ${state.currentSubTaskIndex + 1}/${state.subTasks.length}: ${currentTask} ---`);
    if (!geminiAI) throw new Error("Gemini API key not configured.");

    // --- Create the history context for the prompt ---
    let historyContext = '';
    if (state.planHistory && state.planHistory.length > 0) {
        historyContext = `
          IMPORTANT: You have already generated the following plans in previous steps.
          Your new plan MUST be the logical next step and MUST NOT repeat, overwrite, or contradict this previous work.
          Previous Plans:
          ${JSON.stringify(state.planHistory)}
          `;
    }    


    const promptParts: any[] = [{
        text: `Based on the workbook's current state and the content of the attached file (if any), generate a JSON plan to accomplish ONLY the following sub-task: "${currentTask}"."${historyContext}\n\nOriginal Goal: "${state.userInput}"\nWorkbook Context: ${JSON.stringify(state.workbookContext)}`
    }];
    if (state.pdfFile) {
        const filePart = await uploadPdfToGemini(state.pdfFile);
        promptParts.unshift(filePart);
    }

    const response = await geminiAI.models.generateContent({
        model: "gemini-2.5-pro",
        contents: [{ role: "user", parts: promptParts }],
        config: {
            systemInstruction: [{ text: `Your output MUST be a single JSON object with "plan" and "summary" keys. Schema: ${SYSTEM_PROMPT}` }],
            responseMimeType: "application/json", temperature: 0 
        }
    });

    const result = JSON.parse(response.text ?? '{}');
    const newPlan = result.plan || [];


    return { 
        finalPlans: [...state.finalPlans, result.plan || []],
        planHistory: [...(state.planHistory || []), newPlan],
        planSummary: result.summary || null,
        currentSubTaskIndex: state.currentSubTaskIndex + 1,
        statusUpdates: [`Sub-plan created for: "${currentTask}"`],
    };
};

// =================================================================
// 3. GRAPH EDGES & CONSTRUCTION
// =================================================================

const afterRouter = (state: AgentState) => {
    return state.routingDecision === 'simple' ? 'simplePath' : 'complexPath';
};

const shouldContinueComplex = (state: AgentState) => {
  return state.currentSubTaskIndex < state.subTasks.length ? "continueComplex" : END;
};


const workflow = new StateGraph(AgentStateAnnotation)
  .addNode("router", routerNode)
  .addNode("singlePlanner", singlePlannerNode)
  .addNode("taskAllocator", taskAllocatorNode)
  .addNode("subPlanner", subPlannerNode)

  .addEdge(START, "router")
  .addConditionalEdges("router", afterRouter, {
    simplePath: "singlePlanner",
    complexPath: "taskAllocator",
  })
  .addEdge("singlePlanner", END) // Simple path ends after one plan.
  .addEdge("taskAllocator", "subPlanner")
  .addConditionalEdges("subPlanner", shouldContinueComplex, {
    continueComplex: "subPlanner", // Loop back for the next sub-task.
    [END]: END,
  });

const app = workflow.compile({ checkpointer: new MemorySaver() });

// =================================================================
// 4. NEW PUBLIC API (streaming)
// =================================================================

export async function streamAgenticWorkflow(
    goal: string,
    hints: AgentHints,
    context: any,
    pdfFile: { buffer: Buffer; mimetype: string } | undefined,
    onUpdate: (event: { type: string; data: any }) => void
) {
    if (!geminiAI) {
        throw new Error("Gemini API key not found.");
    }

    const initialState = {
        userInput: goal,
        workbookContext: context,
        hints: hints,
        routingDecision: null,
        subTasks: [],
        currentSubTaskIndex: 0,
        finalPlans: [],
        planHistory: [],
        planSummary: null,
        pdfFile: pdfFile,
        statusUpdates: [],
    };

    const config = { configurable: { thread_id: `agent-run-${Date.now()}` } };
    
    const stream = await app.stream(initialState as any, config);

    for await (const chunk of stream) {
        for (const [nodeName, nodeState] of Object.entries(chunk)) {
            // This block remains the same
            if (nodeState && 'statusUpdates' in nodeState && Array.isArray(nodeState.statusUpdates)) {
                for (const update of nodeState.statusUpdates) {
                    onUpdate({ type: 'status', data: update });
                }
            }
            
            
            if ((nodeName === 'singlePlanner' || nodeName === 'subPlanner') && nodeState) {
                
              
                if ('planSummary' in nodeState && nodeState.planSummary) {
                    onUpdate({ type: 'summary', data: nodeState.planSummary });
                }

                
                if ('finalPlans' in nodeState && Array.isArray(nodeState.finalPlans) && nodeState.finalPlans.length > 0) {
                    const newPlan = nodeState.finalPlans[nodeState.finalPlans.length - 1];
                    if (newPlan) {
                        onUpdate({ type: 'subPlan', data: { plan: newPlan } });
                    }
                }
            }
        }
    }
    
    const finalState = await app.getState(config);
    onUpdate({ 
        type: 'finalResult', 
        data: {
            plans: finalState.values.finalPlans 
        }
    });
}

