import { IDataService } from './IDataService';
import { IProcessStep, parseDependsOn } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';

// Synthetic demo data modeled on a real accounts-payable process flow
// (same steps, decision points, document artifacts, and two-level
// approval structure), but with placeholder names instead of real
// employees/company details, since this file is committed to a public
// repo. Row order matters: DependsOn tokens (e.g. "9.6.1.1-7") are
// resolved by row number (header = row 1), so don't reorder these rows.
const RAW_STEPS: Array<Omit<IProcessStep, 'id' | 'dependsOn'> & { dependsOnRaw: string }> = [
  // 9.6.1.1 - Create Purchase Order
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Send', actionDescription: 'Quote/Order sent to Business', responsibleJobTitle: 'Vendor', shapeOverride: 'Document', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Receive', actionDescription: 'Request for goods or services received', responsibleJobTitle: 'Business', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-2' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Approve', actionDescription: 'Order approved', responsibleJobTitle: 'Business', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Submit', actionDescription: 'Order/Contract sent to Procurement', responsibleJobTitle: 'Business', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-4' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Receive', actionDescription: 'Quote/Order received from Business', responsibleJobTitle: 'Procurement - Demo Region', shapeOverride: 'Document', dependsOnRaw: '9.6.1.1-5' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is vendor set up in the finance system?', responsibleJobTitle: 'Procurement - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.1-6' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up vendor & submit for approval', responsibleJobTitle: 'Procurement - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-7', edgeLabels: { '9.6.1.1-7': 'No' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Non Threshold)', action: 'Approve', actionDescription: 'Vendor approved in finance system', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-8' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Purchase Order created & submitted for approval', responsibleJobTitle: 'Procurement - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-7\n9.6.1.1-9', edgeLabels: { '9.6.1.1-7': 'Yes' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Level 1 approval of Purchase Order (checking amount & coding)', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-10' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is PO value over the high threshold?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.1-11' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Level 2 approval of Purchase Order', responsibleJobTitle: 'Chief Financial Officer - Demo Region 2', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-12', edgeLabels: { '9.6.1.1-12': 'Yes' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'PO approved on finance system', responsibleJobTitle: 'Procurement - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-12\n9.6.1.1-13', edgeLabels: { '9.6.1.1-12': 'No' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Issue', actionDescription: "PO issued to Vendor with Business CC'd", responsibleJobTitle: 'Procurement - Demo Region', shapeOverride: 'Document', dependsOnRaw: '9.6.1.1-14' },

  // 9.6.1.2 - Process Purchase Invoice
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Send', actionDescription: 'Purchase Invoice sent to Business', responsibleJobTitle: 'Vendor', shapeOverride: 'Document', dependsOnRaw: '9.6.1.2-15' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Forward', actionDescription: 'Invoice forwarded to shared AP inbox', responsibleJobTitle: 'Business', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-16' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is invoice covered by a Purchase Order?', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-17' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Process invoice against Purchase Order and send for approval', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-18', edgeLabels: { '9.6.1.2-18': 'Yes' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Process invoice without a Purchase Order and send for approval', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-18', edgeLabels: { '9.6.1.2-18': 'No' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Level 1 approval of invoice (checking amount & coding)', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-19\n9.6.1.2-20' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is invoice value over the high threshold?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-21' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Level 2 approval of invoice', responsibleJobTitle: 'Chief Financial Officer - Demo Region 2', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-22', edgeLabels: { '9.6.1.2-22': 'Yes' } },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Purchase invoice ready for payment', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-22\n9.6.1.2-23', edgeLabels: { '9.6.1.2-22': 'No' } },

  // 9.6.1.3 - Pay Purchase Invoice
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Payment run prepared and sent for approval', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-24' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Payments approved', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-25' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Payments set up on banking platform', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-26' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: '1st level approval of payments on banking platform', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-27' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: '2nd level approval of payments on banking platform', responsibleJobTitle: 'Chief Financial Officer - Demo Region 2', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-28' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Payments posted on finance system', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.3-29' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Issue', actionDescription: 'Proof of payment / remittance sent to vendor', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Document', dependsOnRaw: '9.6.1.3-30' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Execute (Within Limits)', action: 'Reconcile', actionDescription: 'Payments reconciled to bank statement', responsibleJobTitle: 'Finance - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-30' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Purchase Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Bank reconciliation reviewed & approved', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-32' }
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

const MOCK_EMPLOYEES: IEmployee[] = [
  { id: 'e1', name: 'Alex Demo', jobTitle: 'Procurement - Demo Region', region: 'Demo Region' },
  { id: 'e2', name: 'Jordan Sample', jobTitle: 'Finance - Demo Region', region: 'Demo Region' },
  { id: 'e3', name: 'Sam Placeholder', jobTitle: 'Finance Manager - Demo Region', region: 'Demo Region' },
  { id: 'e4', name: 'Riley Example', jobTitle: 'Finance Manager - Demo Region', region: 'Demo Region' },
  { id: 'e5', name: 'Taylor Fixture', jobTitle: 'Chief Financial Officer - Demo Region 2', region: 'Demo Region 2' },
  { id: 'e6', name: 'Morgan Testcase', jobTitle: 'Business', region: 'Demo Region' },
  { id: 'e7', name: 'Casey Stub', jobTitle: 'Finance - Demo Region', region: 'Demo Region' }
];

const MOCK_RISKS: IRiskStatement[] = [
  { id: 'r1', title: 'Segregation of duties - AP', riskStatement: 'Same person creates and approves a purchase order.', linkedProcessStepIds: ['9.6.1.1'] },
  { id: 'r2', title: 'Unauthorized payment', riskStatement: 'Payment released without required approval threshold met.', linkedProcessStepIds: ['9.6.1.3'] }
];

export class MockDataService implements IDataService {
  private _steps: IProcessStep[] = buildMockSteps();

  public getProcessSteps(): Promise<IProcessStep[]> {
    return Promise.resolve(this._steps.slice());
  }

  public getEmployees(): Promise<IEmployee[]> {
    return Promise.resolve(MOCK_EMPLOYEES.slice());
  }

  public getRiskStatements(): Promise<IRiskStatement[]> {
    return Promise.resolve(MOCK_RISKS.slice());
  }

  public addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep> {
    const created: IProcessStep = { ...step, id: `mock-${this._steps.length + 1}` };
    this._steps.push(created);
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
