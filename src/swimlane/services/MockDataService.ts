import { IDataService, IBulkAddStepsResult } from './IDataService';
import { IProcessStep, parseDependsOn } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement, parseLinkedRisks } from '../models/IRiskStatement';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';
import { ICategoryLabel } from '../models/ICategoryLabel';
import { IProcessIdLabel } from '../models/IProcessIdLabel';
import { IProcessIdLock } from '../models/IProcessIdLock';
import { ISwimlaneComment } from '../models/ISwimlaneComment';
import { ISwimlaneStatus, SwimlaneStage } from '../models/ISwimlaneStatus';

// Real accounts-payable process data from the real Master File SharePoint
// list (QLE UK; called "9.6 tester" until 2026-08-19) - used deliberately
// instead of placeholder data, at the user's
// explicit request, so "Use mock data" previews the same real flow while
// SharePoint sign-in is still being debugged. Row order matters: DependsOn
// tokens (e.g. "9.6.1.1-7") are resolved by row number (header = row 1),
// so don't reorder these rows. Yes/No edgeLabels below are inferred from
// which rows share a dependency token (not part of the source data, same
// as the rest of the app - see IProcessStep.edgeLabels).
const RAW_STEPS: Array<Omit<IProcessStep, 'id' | 'dependsOn' | 'linkedRisks'> & { dependsOnRaw: string; linkedRisksRaw?: string }> = [
  // 9.6.1.1 - Create Purchase Order
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Receive', actionDescription: 'Quote/order/Contract received from business', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Document', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is vendor set up in netsuite?', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up vendor on NetSuite & submit for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3', edgeLabels: { '9.6.1.1-3': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Non Threshold)', action: 'Approve', actionDescription: 'CFO approves new vendor set up', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-4', linkedRisksRaw: 'r1:High' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Purchase Order created on NetSuite & routed for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3\n9.6.1.1-5', edgeLabels: { '9.6.1.1-3': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is PO value < $5,000 (or equivalent)', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.1-6' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'NetSuite routes Purchase Order to CFO for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-7', edgeLabels: { '9.6.1.1-7': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Purchase Order on NetSuite', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-8' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Purchase Order approved on NetSuite', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-9\n9.6.1.1-7', edgeLabels: { '9.6.1.1-7': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Issue', actionDescription: "Purchase Order issued to Vendor with Business CC'd", responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Document', dependsOnRaw: '9.6.1.1-10' },

  // 9.6.1.2 - Process Vendor Bill/Invoice
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Submit', actionDescription: 'Vendor Bill/Invoice sent to invoices@qleapenergy.com', responsibleJobTitle: 'Vendor, Business, Finance Manager', shapeOverride: 'Process Step', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill/Invoice covered by a Purchase Order?', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-12', linkedRisksRaw: 'r3:Low' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Create Bill on NetSuite against relevant Purchase Order', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-13', edgeLabels: { '9.6.1.2-13': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'NetSuite completes a 3 way match (3WM) on the Bill against the Purchase order', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-14' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Automated', actionDescription: 'Are there varances, outside of tolerance, on the 3WM?', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-15' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'NetSuite routes Bill to CFO for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-16', edgeLabels: { '9.6.1.2-16': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Bill on NetSuite', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-17' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Bill/invoice is automatically approved on NetSuite', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-16', edgeLabels: { '9.6.1.2-16': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Create Bill on NetSuite  & route for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-13', edgeLabels: { '9.6.1.2-13': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill value < $2,500 (or equivalent)', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-20' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'NetSuite routes Bill to QLE CFO for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-21' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Bill on NetSuite', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-22' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill value < $25,000 (or equivalent)', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-23' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'NetSuite routes Bill to ASPI CFO for approval', responsibleJobTitle: 'Chief Financial Officer - ASPI', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-24', edgeLabels: { '9.6.1.2-24': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Bill on NetSuite', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-22' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Bill approved on NetSuite', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-24\n9.6.1.2-23', edgeLabels: { '9.6.1.2-24': 'Yes' } },

  // 9.6.1.3 - Pay Vendor Bill/Invoice
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Run AP Aging Detail Report showing bills due for payment', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Download to Excel and prepare list of proposed payment.', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Document', dependsOnRaw: '9.6.1.3-28' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Endorse / Recommend', action: 'Recommend', actionDescription: 'Send list of recommended payments to CFO for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-29' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves proposed payment list', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-30', linkedRisksRaw: 'r2:Medium' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Has vendor been set up as a beneficiary on Banking System?', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.3-31' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up Beneficiary on Banking System and send to CFO for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-32', edgeLabels: { '9.6.1.3-32': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Non Threshold)', action: 'Approve', actionDescription: 'CFO approves Beneficiary on Banking System', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-33' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Beneficiary available to pay on Banking System', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-32\n9.6.1.3-34', edgeLabels: { '9.6.1.3-32': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up approved payments on Banking system', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-35' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves payments on Banking System', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-36' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Payments posted on NetSuite against vendor bills', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-37' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Reconcile', actionDescription: 'Payments reconciled to bank statement on NetSuite', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United States', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-38' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Reconcile', actionDescription: 'Bank reconciliation prepared at month-end', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United States', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-39' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO reiews and approves monthly bank rec', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-40' }
];

function buildMockSteps(): IProcessStep[] {
  return RAW_STEPS.map((raw, index) => {
    const { dependsOnRaw, linkedRisksRaw, ...rest } = raw;
    return {
      ...rest,
      id: `mock-${index + 1}`,
      dependsOn: parseDependsOn(dependsOnRaw),
      linkedRisks: parseLinkedRisks(linkedRisksRaw)
    };
  });
}

// Names stay fictional (no real employee data available yet), but job
// titles match the real ResponsibleJobTitle values above so the
// Responsible picker and department filter behave meaningfully against
// them. Shaped like the real "QLE Existing Organisation" list (Job title
// and Department as separate clean fields, no name) rather than the old
// composite-string/fictional-name mock, now that the real schema is
// confirmed.
const MOCK_EMPLOYEES: IEmployee[] = [
  { id: 'e1', jobTitle: 'Finance Manager', department: '6002 - Finance' },
  { id: 'e2', jobTitle: 'Chief Financial Officer', department: '7003 - Finance' },
  { id: 'e3', jobTitle: 'Accounts Payable Analyst', department: '6002 - Finance' },
  { id: 'e4', jobTitle: 'Procurement Manager', department: '6003 - Procurement' }
];

// Shaped like the real "risk register data" list (see the schema comment
// in models/IRiskStatement.ts) - same three AP-relevant risk concepts the
// old mock data used, reshaped into the real columns. Two of these are
// pre-linked to a step each (see linkedRisksRaw above) so the "add a real
// risk to a step" flow has something to look at out of the box in mock
// mode, not just an empty picker.
const MOCK_RISKS: IRiskStatement[] = [
  {
    id: 'r1', riskId: 'OP-014', category: 'Operational / Financial Controls',
    riskStatement: 'Same person creates and approves a purchase order.',
    rootCause: "NetSuite's approval workflow doesn't enforce maker-checker separation for this role.",
    likelihood: 0.4, materiality: 25, inherentRiskRating: 10, riskResponse: 'Mitigate'
  },
  {
    id: 'r2', riskId: 'OP-021', category: 'Operational / Financial Controls',
    riskStatement: 'Payment released without required approval threshold met.',
    rootCause: 'Manual override of banking system approval limits.',
    likelihood: 0.2, materiality: 50, inherentRiskRating: 10, riskResponse: 'Mitigate'
  },
  {
    id: 'r3', riskId: 'OP-033', category: 'Operational / Financial Controls',
    riskStatement: 'Invoice approved without a matching Purchase Order on file.',
    rootCause: '3-way match exception queue not reviewed consistently.',
    likelihood: 0.5, materiality: 10, inherentRiskRating: 5, riskResponse: 'Accept'
  }
];

export class MockDataService implements IDataService {
  private _steps: IProcessStep[] = buildMockSteps();
  // Starts empty - the confirmed real Process Groups already live in
  // apqcHierarchy.ts's static table; this only holds ones a user adds at
  // runtime via "+ Add new process" for a group that table doesn't cover.
  private _groupLabels: IProcessGroupLabel[] = [];
  // Same idea, one level up - names for a genuinely new Category the
  // static APQC_CATEGORY_NAMES table doesn't cover, created via "+ Add
  // new category".
  private _categoryLabels: ICategoryLabel[] = [];
  // Employees added at runtime via "+ Add employee" - kept separate from
  // the static MOCK_EMPLOYEES const rather than mutating it in place,
  // same reasoning as _groupLabels/_categoryLabels layering runtime
  // additions on top of static seed data instead of touching it directly.
  private _addedEmployees: IEmployee[] = [];
  // Same idea, one level down - names for an empty Process ID shell
  // created via "+ Add new process ID" before it has any real steps.
  private _processIdLabels: IProcessIdLabel[] = [];
  // Append-only audit trail - see IProcessIdLock. Starts empty; nothing
  // is locked until someone explicitly locks it.
  private _processIdLocks: IProcessIdLock[] = [];
  // Append-only feedback log - see ISwimlaneComment. Starts empty.
  private _swimlaneComments: ISwimlaneComment[] = [];
  // One mutable record per processId+region, not append-only - see the
  // schema comment on ISwimlaneStatus for why. Starts empty; every
  // swimlane is treated as Draft until someone explicitly sets one.
  private _swimlaneStatuses: ISwimlaneStatus[] = [];
  // A monotonic counter, not `_steps.length + 1` - length-based IDs looked
  // fine until the first delete-then-add in the same session (e.g. undoing
  // a delete): the array shrinks, so the next "length + 1" ID collides
  // with a step that's still there, producing two React children with the
  // same key. Seeded past the real seed data so it never collides with
  // that on first use either.
  private _nextStepId = this._steps.length + 1;

  // Who to stamp new/edited steps' createdBy/modifiedBy with - see the
  // schema comment on IProcessStep for why this mirrors what
  // GraphDataService does for real usage instead of leaving these blank.
  private _currentUserName: string;

  constructor(currentUserName: string) {
    this._currentUserName = currentUserName;
  }

  public getProcessSteps(): Promise<IProcessStep[]> {
    return Promise.resolve(this._steps.slice());
  }

  public getEmployees(): Promise<IEmployee[]> {
    return Promise.resolve([...MOCK_EMPLOYEES, ...this._addedEmployees]);
  }

  public addEmployee(jobTitle: string, department: string): Promise<IEmployee> {
    const created: IEmployee = { id: `mock-employee-${this._addedEmployees.length + 1}`, jobTitle, department: department || undefined };
    this._addedEmployees.push(created);
    return Promise.resolve(created);
  }

  public getRiskStatements(): Promise<IRiskStatement[]> {
    return Promise.resolve(MOCK_RISKS.slice());
  }

  public getCategoryLabels(): Promise<ICategoryLabel[]> {
    return Promise.resolve(this._categoryLabels.slice());
  }

  public addCategoryLabel(categoryId: string, name: string): Promise<ICategoryLabel> {
    const created: ICategoryLabel = { id: `mock-category-${this._categoryLabels.length + 1}`, categoryId, name };
    this._categoryLabels.push(created);
    return Promise.resolve(created);
  }

  public updateCategoryLabel(id: string, name: string): Promise<void> {
    const index = this._categoryLabels.findIndex(l => l.id === id);
    if (index >= 0) this._categoryLabels[index] = { ...this._categoryLabels[index], name };
    return Promise.resolve();
  }

  public getProcessGroupLabels(): Promise<IProcessGroupLabel[]> {
    return Promise.resolve(this._groupLabels.slice());
  }

  public addProcessGroupLabel(groupId: string, name: string): Promise<IProcessGroupLabel> {
    const created: IProcessGroupLabel = { id: `mock-group-${this._groupLabels.length + 1}`, groupId, name };
    this._groupLabels.push(created);
    return Promise.resolve(created);
  }

  public updateProcessGroupLabel(id: string, name: string): Promise<void> {
    const index = this._groupLabels.findIndex(l => l.id === id);
    if (index >= 0) this._groupLabels[index] = { ...this._groupLabels[index], name };
    return Promise.resolve();
  }

  public getProcessIdLabels(): Promise<IProcessIdLabel[]> {
    return Promise.resolve(this._processIdLabels.slice());
  }

  public addProcessIdLabel(processId: string, name: string): Promise<IProcessIdLabel> {
    const created: IProcessIdLabel = { id: `mock-progress-${this._processIdLabels.length + 1}`, processId, name };
    this._processIdLabels.push(created);
    return Promise.resolve(created);
  }

  public updateProcessIdLabel(id: string, name: string): Promise<void> {
    const index = this._processIdLabels.findIndex(l => l.id === id);
    if (index >= 0) this._processIdLabels[index] = { ...this._processIdLabels[index], name };
    return Promise.resolve();
  }

  public getProcessIdLocks(): Promise<IProcessIdLock[]> {
    return Promise.resolve(this._processIdLocks.slice());
  }

  public lockProcessId(processId: string, region: string, lockedBy: string, reason: string): Promise<IProcessIdLock> {
    const created: IProcessIdLock = {
      id: `mock-lock-${this._processIdLocks.length + 1}`,
      processId, region, lockedBy, reason,
      lockedAt: new Date().toISOString(),
      unlockedBy: '', unlockedAt: '', unlockReason: ''
    };
    this._processIdLocks.push(created);
    return Promise.resolve(created);
  }

  public unlockProcessId(id: string, unlockedBy: string, reason: string): Promise<void> {
    const index = this._processIdLocks.findIndex(l => l.id === id);
    if (index >= 0) {
      this._processIdLocks[index] = {
        ...this._processIdLocks[index],
        unlockedBy,
        unlockedAt: new Date().toISOString(),
        unlockReason: reason
      };
    }
    return Promise.resolve();
  }

  public getSwimlaneComments(): Promise<ISwimlaneComment[]> {
    return Promise.resolve(this._swimlaneComments.slice());
  }

  public addSwimlaneComment(processId: string, region: string, author: string, comment: string): Promise<ISwimlaneComment> {
    const created: ISwimlaneComment = {
      id: `mock-comment-${this._swimlaneComments.length + 1}`,
      processId, region, author, comment,
      postedAt: new Date().toISOString()
    };
    this._swimlaneComments.push(created);
    return Promise.resolve(created);
  }

  public getSwimlaneStatuses(): Promise<ISwimlaneStatus[]> {
    return Promise.resolve(this._swimlaneStatuses.slice());
  }

  public addSwimlaneStatus(processId: string, region: string, stage: SwimlaneStage, setBy: string): Promise<ISwimlaneStatus> {
    const created: ISwimlaneStatus = {
      id: `mock-status-${this._swimlaneStatuses.length + 1}`,
      processId, region, stage, setBy,
      setAt: new Date().toISOString()
    };
    this._swimlaneStatuses.push(created);
    return Promise.resolve(created);
  }

  public updateSwimlaneStatus(id: string, stage: SwimlaneStage, setBy: string): Promise<void> {
    const index = this._swimlaneStatuses.findIndex(s => s.id === id);
    if (index >= 0) this._swimlaneStatuses[index] = { ...this._swimlaneStatuses[index], stage, setBy, setAt: new Date().toISOString() };
    return Promise.resolve();
  }

  public addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep> {
    const now = new Date().toISOString();
    const created: IProcessStep = {
      ...step,
      id: `mock-${this._nextStepId++}`,
      createdBy: this._currentUserName, createdAt: now,
      modifiedBy: this._currentUserName, modifiedAt: now
    };
    this._steps.push(created);
    return Promise.resolve(created);
  }

  public addProcessSteps(steps: Array<Omit<IProcessStep, 'id'>>): Promise<IBulkAddStepsResult> {
    const now = new Date().toISOString();
    const created = steps.map(step => {
      const item: IProcessStep = {
        ...step,
        id: `mock-${this._nextStepId++}`,
        createdBy: this._currentUserName, createdAt: now,
        modifiedBy: this._currentUserName, modifiedAt: now
      };
      this._steps.push(item);
      return item;
    });
    // Mock mode has nothing that can fail a single row the way a real
    // network call can - failed always empty, matching IDataService's
    // contract for the real implementation.
    return Promise.resolve({ created, failed: [] });
  }

  public updateProcessStep(step: IProcessStep): Promise<void> {
    const index = this._steps.findIndex(s => s.id === step.id);
    if (index >= 0) {
      this._steps[index] = step;
    }
    return Promise.resolve();
  }

  public deleteProcessStep(id: string): Promise<void> {
    this._steps = this._steps.filter(s => s.id !== id);
    return Promise.resolve();
  }
}
