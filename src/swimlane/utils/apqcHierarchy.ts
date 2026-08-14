// APQC Process Classification Framework (Cross-Industry, v7+) navigation
// levels above Progress ID. Process Step IDs already encode this - e.g.
// "9.6.1.1" is Category 9, Process Group 9.6, Process 9.6.1, Activity
// 9.6.1.1 - these just truncate to fewer segments, same technique
// getProgressId (in models/IProcessStep.ts) already uses for 3 segments.

export function getCategoryId(processStepId: string): string {
  return (processStepId || '').split('.').slice(0, 1).join('.');
}

export function getProcessGroupId(processStepId: string): string {
  return (processStepId || '').split('.').slice(0, 2).join('.');
}

// The 13 top-level categories are a fixed part of the published standard
// (confirmed against APQC's own framework, not guessed) - safe to hard
// code, since this doesn't vary by company or change often.
export const APQC_CATEGORY_NAMES: Record<string, string> = {
  '1': 'Develop Vision and Strategy',
  '2': 'Develop and Manage Products and Services',
  '3': 'Market and Sell Products and Services',
  '4': 'Manage Supply Chain for Physical Products',
  '5': 'Deliver Services',
  '6': 'Manage Customer Service',
  '7': 'Develop and Manage Human Capital',
  '8': 'Manage Information Technology',
  '9': 'Manage Financial Resources',
  '10': 'Acquire, Construct, and Manage Assets',
  '11': 'Manage Enterprise Risk, Compliance, Remediation, and Resiliency',
  '12': 'Manage External Relationships',
  '13': 'Develop and Manage Business Capabilities'
};

// Process Group names - NOT the complete APQC list (each category has
// many groups; the full set is a much longer reference than is useful to
// hard-code up front) - just the ones confirmed so far. Add more here as
// new Process Step ID prefixes show up in real data.
export const APQC_PROCESS_GROUP_NAMES: Record<string, string> = {
  '9.6': 'Process Accounts Payable and Expense Reimbursements'
};

export function getCategoryName(categoryId: string): string {
  return APQC_CATEGORY_NAMES[categoryId] || `Category ${categoryId}`;
}

export function getProcessGroupName(groupId: string): string {
  return APQC_PROCESS_GROUP_NAMES[groupId] || `Process Group ${groupId}`;
}
