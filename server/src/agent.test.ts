import { streamAgenticWorkflow } from './agent.js';
import { GoogleGenAI } from '@google/genai';

// --- Mock Setup ---
jest.mock('@google/genai', () => {
  const mockGenerateContent = jest.fn();
  // We also need to mock the file upload function now.
  const mockUploadFile = jest.fn().mockResolvedValue({
    uri: 'mock-file-uri',
    mimeType: 'application/pdf',
  });
  
  return {
    GoogleGenAI: jest.fn().mockImplementation(() => ({
      models: { generateContent: mockGenerateContent },
      files: { upload: mockUploadFile },
    })),
  };
});

const mockGeminiAI = new GoogleGenAI({});
const mockGenerateContent = mockGeminiAI.models.generateContent as jest.Mock;

describe('Agentic Workflow with PDF Handling', () => {

  beforeEach(() => {
    mockGenerateContent.mockClear();
  });

  it('should follow the simple path and send the PDF to the planner', async () => {
    // Arrange: Mock responses for the router and the single planner.
    mockGenerateContent
      .mockResolvedValueOnce({ text: JSON.stringify({ decision: 'simple' }) })
      .mockResolvedValueOnce({ text: JSON.stringify({ plan: [{ op: 'plan_from_pdf' }], summary: 'Simple plan summary.' }) });

    const mockPdfFile = { buffer: Buffer.from('test pdf'), mimetype: 'application/pdf' };
    const onUpdate = jest.fn(); // Use a Jest mock function to easily track calls.

    // Act: Run the workflow with a PDF file.
    await streamAgenticWorkflow("Summarize the attached PDF", {}, { sheets: [] }, mockPdfFile, onUpdate);

    // Assert: Verify the API was called correctly.
    expect(mockGenerateContent).toHaveBeenCalledTimes(2);

    // Call 1: Router (should NOT have file data).
    const routerCallArgs = mockGenerateContent.mock.calls[0][0];
    expect(routerCallArgs.contents[0].parts.some((p: any) => p.fileData)).toBe(false);

    // Call 2: Single Planner (SHOULD have file data).
    const plannerCallArgs = mockGenerateContent.mock.calls[1][0];
    expect(plannerCallArgs.contents[0].parts.some((p: any) => p.fileData)).toBe(true);
    
    // Check that the final plan was received.
    const finalResultEvent = onUpdate.mock.calls.find(call => call[0].type === 'finalResult');
    expect(finalResultEvent[0].data.plans[0][0].op).toBe('plan_from_pdf');
  });

  it('should follow the complex path and send the PDF to the allocator and sub-planners', async () => {
    // Arrange: Mock a sequence of responses for the complex workflow.
    mockGenerateContent
      .mockResolvedValueOnce({ text: JSON.stringify({ decision: 'complex' }) }) // 1. Router
      .mockResolvedValueOnce({ text: JSON.stringify({ subTasks: ['Analyze intro', 'Analyze conclusion'] }) }) // 2. Allocator
      .mockResolvedValueOnce({ text: JSON.stringify({ plan: [{ op: 'plan_for_intro' }], summary: 'Intro plan' }) }) // 3. Sub-plan 1
      .mockResolvedValueOnce({ text: JSON.stringify({ plan: [{ op: 'plan_for_conclusion' }], summary: 'Conclusion plan' }) }); // 4. Sub-plan 2

    const mockPdfFile = { buffer: Buffer.from('test pdf'), mimetype: 'application/pdf' };
    const onUpdate = jest.fn();

    // Act
    await streamAgenticWorkflow("Analyze the intro and conclusion of the attached PDF", {}, { sheets: [] }, mockPdfFile, onUpdate);
    
    // Assert
    expect(mockGenerateContent).toHaveBeenCalledTimes(4);

    // Call 1: Router (no PDF).
    const routerArgs = mockGenerateContent.mock.calls[0][0];
    expect(routerArgs.contents[0].parts.some((p: any) => p.fileData)).toBe(false);

    // Call 2: Task Allocator (should have PDF).
    const allocatorArgs = mockGenerateContent.mock.calls[1][0];
    expect(allocatorArgs.contents[0].parts.some((p: any) => p.fileData)).toBe(true);
    
    // Call 3: Sub-Planner 1 (should have PDF).
    const subPlanner1Args = mockGenerateContent.mock.calls[2][0];
    expect(subPlanner1Args.contents[0].parts.some((p: any) => p.fileData)).toBe(true);

    // Call 4: Sub-Planner 2 (should have PDF).
    const subPlanner2Args = mockGenerateContent.mock.calls[3][0];
    expect(subPlanner2Args.contents[0].parts.some((p: any) => p.fileData)).toBe(true);
    
    // Check final result
    const finalResultEvent = onUpdate.mock.calls.find(call => call[0].type === 'finalResult');
    expect(finalResultEvent[0].data.plans.length).toBe(2);
    expect(finalResultEvent[0].data.plans[0][0].op).toBe('plan_for_intro');
    expect(finalResultEvent[0].data.plans[1][0].op).toBe('plan_for_conclusion');
  });
});

// --- NEW TEST SUITE FOR PLAN HISTORY LOGIC ---
describe('Agentic Workflow with Plan History', () => {
  beforeEach(() => {
    mockGenerateContent.mockClear();
  });

  it('should pass the plan from the first sub-planner as history to the second sub-planner', async () => {
    // Arrange: Mock the full complex path with specific plans.
    const firstPlan = [{ op: 'first_step', detail: 'Analyze the income statement' }];
    const secondPlan = [{ op: 'second_step', detail: 'Analyze the balance sheet' }];

    mockGenerateContent
      .mockResolvedValueOnce({ text: JSON.stringify({ decision: 'complex' }) }) // 1. Router
      .mockResolvedValueOnce({ text: JSON.stringify({ subTasks: ['Task 1', 'Task 2'] }) }) // 2. Allocator
      .mockResolvedValueOnce({ text: JSON.stringify({ plan: firstPlan, summary: 'Summary 1' }) }) // 3. Sub-plan 1
      .mockResolvedValueOnce({ text: JSON.stringify({ plan: secondPlan, summary: 'Summary 2' }) }); // 4. Sub-plan 2

    const onUpdate = jest.fn();

    // Act
    await streamAgenticWorkflow("Complex task", {}, { sheets: [] }, undefined, onUpdate);

    // Assert
    expect(mockGenerateContent).toHaveBeenCalledTimes(4);

    // Call 3: Sub-Planner 1 (should NOT have history).
    const subPlanner1Args = mockGenerateContent.mock.calls[2][0];
    const prompt1 = subPlanner1Args.contents[0].parts[0].text;
    expect(prompt1).not.toContain('Previous Plans:');
    expect(prompt1).not.toContain('IMPORTANT: You have already generated');

    // Call 4: Sub-Planner 2 (SHOULD have history).
    const subPlanner2Args = mockGenerateContent.mock.calls[3][0];
    const prompt2 = subPlanner2Args.contents[0].parts[0].text;
    expect(prompt2).toContain('IMPORTANT: You have already generated');
    expect(prompt2).toContain('Previous Plans:');
    
    // Verify that the first plan is included in the prompt for the second planner.
    // The history is an array of plans, so it will be [[...firstPlan]].
    expect(prompt2).toContain(JSON.stringify([firstPlan]));

    // Check final result just to be sure
    const finalResultEvent = onUpdate.mock.calls.find(call => call[0].type === 'finalResult');
    expect(finalResultEvent[0].data.plans.length).toBe(2);
    expect(finalResultEvent[0].data.plans[0][0].op).toBe('first_step');
    expect(finalResultEvent[0].data.plans[1][0].op).toBe('second_step');
  });
});