import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { IEmployee } from '../models/IEmployee';
import { AuthorityTier, getAuthorityTier } from '../models/IProcessStep';
import { matchesAuthorityTier } from '../utils/authoritySuggestion';

export interface IEmployeePickerProps {
  employees: IEmployee[];
  actionType: string; // used to derive the authority tier for the suggested-first ordering
  value: string | undefined; // ResponsibleJobTitle currently on the task
  onChange: (jobTitle: string) => void;
}

// Lanes are always job titles, never a person's name - and by explicit
// user choice, employee names aren't pulled into the app at all (see
// models/IEmployee.ts), so several people holding the same title collapse
// into one option here rather than one row per person. Used to also
// accept a department filter to narrow this list, but that turned out to
// be genuinely confusing in practice - it sat in the same toolbar as
// filters that visibly change the diagram (Region, Process Step tabs),
// while this one only ever affected this one dropdown, several screens
// away, with no visible feedback that anything had happened. Removed at
// a real user's request rather than kept "just in case".
const EmployeePicker: React.FC<IEmployeePickerProps> = ({ employees, actionType, value, onChange }) => {
  const tier: AuthorityTier = getAuthorityTier(actionType);

  // Suggested-first ordering only, per the confirmed rule that AI/job-title
  // suggestions should differentiate by authority tier (Execute -> analyst,
  // Endorse -> manager, Approve -> senior/chief). This is a heuristic
  // ordering, not a real AI call - see utils/authoritySuggestion.ts for why.
  const options: IDropdownOption[] = React.useMemo(() => {
    const distinctTitles = Array.from(new Set(employees.map(e => e.jobTitle)));
    const sorted = distinctTitles.slice().sort((a, b) => {
      const aMatches = matchesAuthorityTier(a, tier) ? 0 : 1;
      const bMatches = matchesAuthorityTier(b, tier) ? 0 : 1;
      return aMatches - bMatches;
    });
    return sorted.map(jobTitle => ({ key: jobTitle, text: jobTitle }));
  }, [employees, tier]);

  return (
    <Dropdown
      label="Responsible"
      placeholder="Select a job title"
      selectedKey={value}
      options={options}
      onChange={(_e, option?: IDropdownOption) => {
        if (option) {
          onChange(String(option.key));
        }
      }}
    />
  );
};

export default EmployeePicker;
