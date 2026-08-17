import { IDataService } from './IDataService';
import { IProcessStep, parseDependsOn } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';

// Real accounts-payable process data from the "9.6 tester" SharePoint list
// (QLE UK) - used deliberately instead of placeholder data, at the user's
// explicit request, so "Use mock data" previews the same real flow while
// SharePoint sign-in is still being debugged. Row order matters: DependsOn
// tokens (e.g. "9.6.1.1-7") are resolved by row number (header = row 1),
// so don't reorder these rows. Yes/No edgeLabels below are inferred from
// which rows share a dependency token (not part of the source data, same
// as the rest of the app - see IProcessStep.edgeLabels).
const RAW_STEPS: Array<Omit<IProcessStep, 'id' | 'dependsOn'> & { dependsOnRaw: string }> = [
  // 9.6.1.1 - Create Purchase Order
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Receive', actionDescription: 'Quote/order/Contract received from business', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Document', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is vendor set up in netsuite?', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up vendor on NetSuite & submit for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3', edgeLabels: { '9.6.1.1-3': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Non Threshold)', action: 'Approve', actionDescription: 'CFO approves new vendor set up', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-4' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Purchase Order created on NetSuite & routed for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3\n9.6.1.1-5', edgeLabels: { '9.6.1.1-3': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is PO value < $5,000 (or equivalent)', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.1-6' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'NetSuite routes Purchase Order to CFO for approval', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-7', edgeLabels: { '9.6.1.1-7': 'No' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Purchase Order on NetSuite', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-8' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Purchase Order approved on NetSuite', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-9\n9.6.1.1-7', edgeLabels: { '9.6.1.1-7': 'Yes' } },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Issue', actionDescription: "Purchase Order issued to Vendor with Business CC'd", responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Document', dependsOnRaw: '9.6.1.1-10' },

  // 9.6.1.2 - Process Vendor Bill/Invoice
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Submit', actionDescription: 'Vendor Bill/Invoice sent to invoices@qleapenergy.com', responsibleJobTitle: 'Vendor, Business, Finance Manager', shapeOverride: 'Process Step', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill/Invoice covered by a Purchase Order?', responsibleJobTitle: 'Finance Manager / 6002 - Finance / United Kingdom', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-12' },
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
  { apqcTitle: '9.6.1 - UK', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves proposed payment list', responsibleJobTitle: 'Chief Financial Officer / 7003 - Finance / United States', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-30' },
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
    const { dependsOnRaw, ...rest } = raw;
    return {
      ...rest,
      id: `mock-${index + 1}`,
      dependsOn: parseDependsOn(dependsOnRaw)
    };
  });
}

// Names stay fictional (no real employee data available yet), but job
// titles match the real ResponsibleJobTitle values above so the
// Responsible picker and region filter behave meaningfully against them.
// Shaped like the real "QLE Existing Organisation" list (Job title and
// Department as separate clean fields, no name) rather than the old
// composite-string/fictional-name mock, now that the real schema is
// confirmed.
const MOCK_EMPLOYEES: IEmployee[] = [
  { id: 'e1', jobTitle: 'Finance Manager', region: '6002 - Finance' },
  { id: 'e2', jobTitle: 'Chief Financial Officer', region: '7003 - Finance' },
  { id: 'e3', jobTitle: 'Accounts Payable Analyst', region: '6002 - Finance' },
  { id: 'e4', jobTitle: 'Procurement Manager', region: '6003 - Procurement' }
];

const MOCK_RISKS: IRiskStatement[] = [
  { id: 'r1', title: 'Segregation of duties - AP', riskStatement: 'Same person creates and approves a purchase order.', linkedProcessStepIds: ['9.6.1.1'], riskLevel: 'High' },
  { id: 'r2', title: 'Unauthorized payment', riskStatement: 'Payment released without required approval threshold met.', linkedProcessStepIds: ['9.6.1.3'], riskLevel: 'Medium' },
  { id: 'r3', title: 'Invoice processed without PO', riskStatement: 'Invoice approved without a matching Purchase Order on file.', linkedProcessStepIds: ['9.6.1.2'], riskLevel: 'Low' }
];

export class MockDataService implements IDataService {
  private _steps: IProcessStep[] = buildMockSteps();
  // Starts empty - the confirmed real Process Groups already live in
  // apqcHierarchy.ts's static table; this only holds ones a user adds at
  // runtime via "+ Add new process" for a group that table doesn't cover.
  private _groupLabels: IProcessGroupLabel[] = [];

  public getProcessSteps(): Promise<IProcessStep[]> {
    return Promise.resolve(this._steps.slice());
  }

  public getEmployees(): Promise<IEmployee[]> {
    return Promise.resolve(MOCK_EMPLOYEES.slice());
  }

  public getRiskStatements(): Promise<IRiskStatement[]> {
    return Promise.resolve(MOCK_RISKS.slice());
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

  public addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep> {
    const created: IProcessStep = { ...step, id: `mock-${this._steps.length + 1}` };
    this._steps.push(created);
    return Promise.resolve(created);
  }

  public addProcessSteps(steps: Array<Omit<IProcessStep, 'id'>>): Promise<IProcessStep[]> {
    const created = steps.map(step => {
      const item: IProcessStep = { ...step, id: `mock-${this._steps.length + 1}` };
      this._steps.push(item);
      return item;
    });
    return Promise.resolve(created);
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
