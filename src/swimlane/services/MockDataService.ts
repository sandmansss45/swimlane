import { IDataService } from './IDataService';
import { IProcessStep, parseDependsOn } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';

// Synthetic demo data - same shape/structure as a real accounts-payable
// import (same row count, dependency tokens, decision/approval mix,
// multi-region lanes) so every feature (drill-down, dependency arrows,
// region filter, authority-tier suggestion) still demos correctly, but no
// real employee names, cost-center codes, or approval thresholds. This
// file is committed to a public repo - keep it that way; if you need to
// test against the real dataset locally, do it in an untracked file, not
// here. Row order still matters: DependsOn tokens (e.g. "9.6.1.1-3") are
// resolved by row number (header = row 1), so don't reorder these rows.
const RAW_STEPS: Array<Omit<IProcessStep, 'id' | 'dependsOn'> & { dependsOnRaw: string }> = [
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Receive', actionDescription: 'Quote/order/Contract received from business', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Document', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is vendor set up in the finance system?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up vendor & submit for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Non Threshold)', action: 'Approve', actionDescription: 'CFO approves new vendor set up', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-4' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Purchase Order created & routed for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.1-3\n9.6.1.1-5' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is PO value under the low threshold?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.1-6' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'System routes Purchase Order to CFO for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-7' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Purchase Order', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.1-8' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Purchase Order approved in the system', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.1-9\n9.6.1.1-7' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.1', processStepName: 'Create Purchase Order', actionType: 'Execute (Within Limits)', action: 'Issue', actionDescription: "Purchase Order issued to Vendor with Business CC'd", responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Document', dependsOnRaw: '9.6.1.1-10' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Submit', actionDescription: 'Vendor Bill/Invoice submitted to shared inbox', responsibleJobTitle: 'Vendor, Business, Finance Manager', shapeOverride: 'Process Step', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill/Invoice covered by a Purchase Order?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-12' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Create Bill against relevant Purchase Order', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-13' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'System completes a 3-way match on the Bill against the Purchase Order', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-14' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Automated', actionDescription: 'Are there variances, outside of tolerance, on the match?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-15' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'System routes Bill to CFO for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-16' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Bill', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-17' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'Bill/invoice is automatically approved', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-16' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Create Bill & route for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.2-13' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill value under the low threshold?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-20' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'System routes Bill to Group CFO for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-21' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Bill', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-22' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Is Bill value under the high threshold?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.2-23' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'System routes Bill to Regional CFO for approval', responsibleJobTitle: 'Chief Financial Officer - Demo Region 2', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-24' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves Bill', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.2-22' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.2', processStepName: 'Process Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Automated', actionDescription: 'Bill approved', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'System workflow', dependsOnRaw: '9.6.1.2-24\n9.6.1.2-23' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Run Aging Detail Report showing bills due for payment', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Prepare list of proposed payments.', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Document', dependsOnRaw: '9.6.1.3-28' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Endorse / Recommend', action: 'Recommend', actionDescription: 'Send list of recommended payments to CFO for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-29' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves proposed payment list', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-30' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Has vendor been set up as a beneficiary on the banking system?', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Decision', dependsOnRaw: '9.6.1.3-31' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up Beneficiary on banking system and send to CFO for approval', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-32' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Non Threshold)', action: 'Approve', actionDescription: 'CFO approves Beneficiary on banking system', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-33' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Review', actionDescription: 'Beneficiary available to pay on banking system', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-32\n9.6.1.3-34' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Set up approved payments on banking system', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-35' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO approves payments on banking system', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Approval', dependsOnRaw: '9.6.1.3-36' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Create', actionDescription: 'Payments posted against vendor bills', responsibleJobTitle: 'Finance Manager - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-37' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Reconcile', actionDescription: 'Payments reconciled to bank statement', responsibleJobTitle: 'Finance Manager - Demo Region 2', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-38' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Execute (Within Limits)', action: 'Reconcile', actionDescription: 'Bank reconciliation prepared at month-end', responsibleJobTitle: 'Finance Manager - Demo Region 2', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-39' },
  { apqcTitle: '9.6.1 - Demo Region', processDescription: 'Process accounts payable', processStepId: '9.6.1.3', processStepName: 'Pay Vendor Bill/Invoice', actionType: 'Approve (Within Thresholds)', action: 'Approve', actionDescription: 'CFO reviews and approves monthly bank reconciliation', responsibleJobTitle: 'Chief Financial Officer - Demo Region', shapeOverride: 'Process Step', dependsOnRaw: '9.6.1.3-40' }
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
  { id: 'e1', name: 'Alex Demo', jobTitle: 'Finance Manager - Demo Region', region: 'Demo Region' },
  { id: 'e2', name: 'Jordan Sample', jobTitle: 'Finance Manager - Demo Region', region: 'Demo Region' },
  { id: 'e3', name: 'Sam Placeholder', jobTitle: 'Finance Manager - Demo Region 2', region: 'Demo Region 2' },
  { id: 'e4', name: 'Riley Example', jobTitle: 'Procurement - Demo Region', region: 'Demo Region' },
  { id: 'e5', name: 'Taylor Fixture', jobTitle: 'Chief Financial Officer - Demo Region', region: 'Demo Region 2' },
  { id: 'e6', name: 'Morgan Testcase', jobTitle: 'Finance Manager - Demo Region', region: 'Demo Region' },
  { id: 'e7', name: 'Casey Stub', jobTitle: 'Controller - Demo Region 2', region: 'Demo Region 2' }
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
