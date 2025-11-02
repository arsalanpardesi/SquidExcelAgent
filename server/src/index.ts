import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import multer from 'multer';
import * as XLSX from 'xlsx';

import { SheetModel } from './sheet.js';
import { streamPlanAndExecute, planAndExecuteForAddin, AgentStreamEvent } from './agent.js'; //should delete this later on
import { extractTextFromPdf } from './pdf.js';
import { parseTenKToStructured } from './parser.js';
import { streamAgenticWorkflow, AgentHints } from './agent.js'; 

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 3001;

const app = express();
app.use(cors());
app.use(express.json({ limit: '16mb' }));

const upload = multer({ storage: multer.memoryStorage() });
const model = new SheetModel();

/* ================= Workbook APIs ================= */

app.get('/api/workbook', (_req, res) => res.json(model.toJSON()));

app.post('/api/sheetOps', (req, res) => {
  const { op, args } = req.body || {};
  try {
    model.dispatch(op, args);
    model.evaluateAll();
    res.json({ ok: true, workbook: model.toJSON() });
  } catch (e) {
    res.status(400).json({ ok: false, error: (e as Error).message });
  }
});

app.post('/api/undo', (_req, res) => {
  const ev = model.undo();
  model.evaluateAll();
  res.json({ undone: ev?.summary ?? null, workbook: model.toJSON() });
});

app.get('/api/provenance', (req, res) => {
  const sheet = String(req.query.sheet || 'Sheet1');
  const cell = String(req.query.cell || 'A1');
  try {
    const prov = model.getProvenance(sheet, cell);
    res.json({ sheet, cell, provenance: prov });
  } catch (e) {
    res.status(400).json({ error: (e as Error).message });
  }
});

/* ========== Import/Ingest Endpoints ========== */
app.post('/api/ingest-10k', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, error: 'No PDF uploaded' });

    const sourceLabel = req.file.originalname || '10-K.pdf';
    const text = await extractTextFromPdf(req.file.buffer);
    if (!text.trim()) return res.status(400).json({ ok: false, error: 'Could not extract text from PDF' });

    const parsed = await parseTenKToStructured(text, sourceLabel);

    const headerIncome = ['Line item', ...parsed.income.periods];
    try { (model as any).deleteSheet('P&L'); } catch {}
    (model as any).createSheet('P&L');
    (model as any).setValues({ sheet: 'P&L', r1: 0, c1: 0, r2: 0, c2: headerIncome.length - 1 }, [headerIncome]);
    parsed.income.lines.forEach((l, i) =>
      (model as any).setValues({ sheet: 'P&L', r1: i+1, c1: 0, r2: i+1, c2: headerIncome.length - 1 }, [[ l.name, ...l.values ]])
    );

    const headerBS = ['Line item', ...parsed.balance.periods];
    try { (model as any).deleteSheet('Balance Sheet'); } catch {}
    (model as any).createSheet('Balance Sheet');
    (model as any).setValues({ sheet: 'Balance Sheet', r1: 0, c1: 0, r2: 0, c2: headerBS.length - 1 }, [headerBS]);
    parsed.balance.lines.forEach((l, i) =>
      (model as any).setValues({ sheet: 'Balance Sheet', r1: i+1, c1: 0, r2: i+1, c2: headerBS.length - 1 }, [[ l.name, ...l.values ]])
    );

    const headerCF = ['Line item', ...parsed.cashflow.periods];
    try { (model as any).deleteSheet('Cash Flow'); } catch {}
    (model as any).createSheet('Cash Flow');
    (model as any).setValues({ sheet: 'Cash Flow', r1: 0, c1: 0, r2: 0, c2: headerCF.length - 1 }, [headerCF]);
    parsed.cashflow.lines.forEach((l, i) =>
      (model as any).setValues({ sheet: 'Cash Flow', r1: i+1, c1: 0, r2: i+1, c2: headerCF.length - 1 }, [[ l.name, ...l.values ]])
    );

    res.json({ ok: true, parsed, workbook: model.toJSON() });
  } catch (e) {
    res.status(500).json({ ok: false, error: (e as Error).message });
  }
});

app.post('/api/import-xlsx', (req, res) => {
  try {
    const workbookData = req.body;
    if (!workbookData || !Array.isArray(workbookData.sheets)) {
      throw new Error('Invalid workbook data format.');
    }
    model.loadFromJSON(workbookData);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ ok: false, error: (e as Error).message });
  }
});


/* ================= Export workbook as .xlsx ================= */
app.get('/api/export-xlsx', (_req, res) => {
  try {
    const book = XLSX.utils.book_new();
    const wb = (model as any).wb as { sheets: Map<string, any> };
    for (const [name, s] of wb.sheets.entries()) {
      const aoa: any[][] = s.rows.map((row: any[]) =>
        row.map((c: any) => (c?.formula ? c.formula : (c?.value ?? null)))
      );
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(book, ws, name.slice(0, 31));
    }
    const buffer = XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="export.xlsx"');
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ ok: false, error: (e as Error).message });
  }
});

/* ================= Agent endpoints ================= */


// This endpoint is for the Excel add-in, which sends its own context.
app.post('/api/agent', upload.single('pdfFile'), async (req, res) => {
  try {
    const { goal, context, sheetHint, insertRow, pdfFileMimeType } = req.body;
    
    if (!goal || !context) {
      // Send a standard error response if initial validation fails
      return res.status(400).json({ error: 'Missing goal or context from add-in.' });
    }

    // 1. Set headers for a Server-Sent Events (SSE) stream
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders(); // Establish the connection immediately

    // 2. Create a helper function to send formatted events to the client
    const sendEvent = (event: { type: string, data: any }) => {
      res.write(`event: ${event.type}\n`);
      res.write(`data: ${JSON.stringify(event.data)}\n\n`);
    };

    const hints: AgentHints = {
      sheetHint,
      insertRow: insertRow ? Number(insertRow) : undefined
    };

    let pdfFile;
    if (req.file) {
      pdfFile = {
        buffer: req.file.buffer,
        mimetype: req.file.mimetype || pdfFileMimeType,
      };
    }

    // 3. Call the new streaming workflow, passing the sendEvent function as the callback
    await streamAgenticWorkflow(goal, hints, JSON.parse(context), pdfFile, sendEvent);
    
    // 4. Signal that the stream has finished successfully
    sendEvent({ type: 'done', data: 'Stream finished.' });
    res.end();

  } catch (e) {
    console.error("Error in streaming /api/agent endpoint:", e);
    // If an error occurs during the process, send an error event before closing
    res.write(`event: error\n`);
    res.write(`data: ${JSON.stringify({ message: (e as Error).message })}\n\n`);
    res.end();
  }
});


/* ================= Static client ================= */
app.use('/', express.static(path.join(__dirname, '..', 'public')));

app.listen(PORT, () => {
  console.log(`Squid AI agent server listening on http://localhost:${PORT}`);
  if (!process.env.GEMINI_API_KEY) {
    console.warn('⚠️  GEMINI_API_KEY not found in .env file. The Gemini model will not be available.');
  }
});