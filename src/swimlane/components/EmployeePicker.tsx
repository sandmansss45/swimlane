import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { IEmployee, formatEmployeeLabel } from '../models/IEmployee';
import { AuthorityTier, getAuthorityTier } from '../models/IProcessStep';
import { matchesAuthorityTier } from '../utils/authoritySuggestion';

export interface IEmployeePickerProps {
  employees: IEmployee[];
  selectedRegion: string | undefined; // from RegionFilter - narrows the list before showing options
  actionType: string; // used to derive the authority tier for the suggested-first ordering
  value: string | undefined; // ResponsibleJobTitle currently on the task
  onChange: (jobTitle: string) => void;
}

// Lanes are always job titles, never a person's name - the picker itself
// still shows "Title — Name" so a human can tell people with the same
// title apart, but onChange only ever passes back the job title.
const EmployeePicker: React.FC<IEmployeePickerProps> = ({ employees, selectedRegion, actionType, value, onChange }) => {
  const tier: AuthorityTier = getAuthorityTier(actionType);

  const narrowed = React.useMemo(
    () => (selectedRegion ? employees.filter(e => e.region === selectedRegion) : employees),
    [employees, selectedRegion]
  );

  // Suggested-first ordering only, per the confirmed rule that AI/job-title
  // suggestions should differentiate by authority tier (Execute -> analyst,
  // Endorse -> manager, Approve -> senior/chief). This is a heuristic
  // ordering, not a real AI call - see utils/authoritySuggestion.ts for why.
  const options: IDropdownOption[] = React.useMemo(() => {
    const sorted = narrowed.slice().sort((a, b) => {
      const aMatches = matchesAuthorityTier(a.jobTitle, tier) ? 0 : 1;
      const bMatches = matchesAuthorityTier(b.jobTitle, tier) ? 0 : 1;
      return aMatches - bMatches;
    });
    return sorted.map(e => ({ key: e.jobTitle, text: formatEmployeeLabel(e) }));
  }, [narrowed, tier]);

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
