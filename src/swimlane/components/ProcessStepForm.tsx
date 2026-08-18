import * as React from 'react';
import { TextField, Dropdown, IDropdownOption, ComboBox, IComboBoxOption } from '@fluentui/react';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement, IRiskLink } from '../models/IRiskStatement';
import EmployeePicker from './EmployeePicker';
import RiskLinkPicker from './RiskLinkPicker';

// The full set of fields a process step actually has - shared by the
// "Add a step" card and the shape edit panel so a step created here has
// every field a step edited there does, and the two forms can never
// silently drift out of sync with each other.
export interface IProcessStepFormValue {
  action: string;
  actionDescription: string;
  actionType: string;
  shapeOverride: string;
  responsibleJobTitle: string;
  dependsOnStepIds: string[];
  linkedRisks: IRiskLink[];
}

export const SHAPE_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(use Action Type / wording heuristic)' },
  { key: 'Process Step', text: 'Process (rounded rectangle)' },
  { key: 'Decision', text: 'Decision (diamond)' },
  { key: 'Approval', text: 'Approval (circle)' },
  { key: 'Document', text: 'Document (artifact)' }
];

// Both Action and Action Type are free text in the real data (there's no
// closed, confirmed list of every value a real export might use) - these
// are just the ones seen so far, offered as suggestions via a freeform
// combo box rather than a closed dropdown, so an unfamiliar existing
// value still displays correctly instead of showing blank.
const ACTION_TYPE_SUGGESTIONS: IComboBoxOption[] = [
  'Execute (Within Limits)', 'Execute (Non Threshold)', 'Approve (Within Thresholds)',
  'Approve (Non Threshold)', 'Endorse / Recommend', 'Automated'
].map(v => ({ key: v, text: v }));

const ACTION_SUGGESTIONS: IComboBoxOption[] = [
  'Send', 'Receive', 'Forward', 'Review', 'Create', 'Submit', 'Approve', 'Automated', 'Issue', 'Reconcile', 'Recommend'
].map(v => ({ key: v, text: v }));

export interface IProcessStepFormProps {
  value: IProcessStepFormValue;
  onChange: (value: IProcessStepFormValue) => void;
  employees: IEmployee[];
  dependsOnOptions: IDropdownOption[];
  riskStatements: IRiskStatement[];
}

const ProcessStepForm: React.FC<IProcessStepFormProps> = ({ value, onChange, employees, dependsOnOptions, riskStatements }) => {
  const set = <K extends keyof IProcessStepFormValue>(key: K, v: IProcessStepFormValue[K]): void => {
    onChange({ ...value, [key]: v });
  };

  return (
    <>
      <TextField
        label="Action description"
        multiline
        value={value.actionDescription}
        onChange={(_e, v) => set('actionDescription', v || '')}
      />
      <ComboBox
        label="Action"
        placeholder="Choose from the list, or type your own..."
        text={value.action}
        allowFreeform
        autoComplete="on"
        options={ACTION_SUGGESTIONS}
        onChange={(_e, option, _index, freeformValue) => set('action', option ? String(option.key) : (freeformValue || ''))}
      />
      <ComboBox
        label="Action type"
        placeholder="Choose from the list, or type your own..."
        text={value.actionType}
        allowFreeform
        autoComplete="on"
        options={ACTION_TYPE_SUGGESTIONS}
        onChange={(_e, option, _index, freeformValue) => set('actionType', option ? String(option.key) : (freeformValue || ''))}
      />
      <Dropdown
        label="Shape"
        selectedKey={value.shapeOverride}
        options={SHAPE_OPTIONS}
        onChange={(_e, option) => option && set('shapeOverride', String(option.key))}
      />
      <EmployeePicker
        employees={employees}
        value={value.responsibleJobTitle}
        onChange={jobTitle => set('responsibleJobTitle', jobTitle)}
      />
      <Dropdown
        label="Depends on"
        placeholder="Which step(s) does this follow?"
        multiSelect
        selectedKeys={value.dependsOnStepIds}
        options={dependsOnOptions}
        onChange={(_e, option) => {
          if (!option) return;
          const ids = option.selected
            ? [...value.dependsOnStepIds, String(option.key)]
            : value.dependsOnStepIds.filter(id => id !== option.key);
          set('dependsOnStepIds', ids);
        }}
      />
      <RiskLinkPicker
        riskStatements={riskStatements}
        value={value.linkedRisks}
        onChange={links => set('linkedRisks', links)}
      />
    </>
  );
};

export default ProcessStepForm;
