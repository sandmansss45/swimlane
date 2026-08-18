import * as React from 'react';
import { ComboBox, IComboBoxOption } from '@fluentui/react';
import { IEmployee } from '../models/IEmployee';

export interface IEmployeePickerProps {
  employees: IEmployee[];
  value: string | undefined; // ResponsibleJobTitle currently on the task
  onChange: (jobTitle: string) => void;
}

// Lanes are always job titles, never a person's name - and by explicit
// user choice, employee names aren't pulled into the app at all (see
// models/IEmployee.ts), so several people holding the same title collapse
// into one option here rather than one row per person. Search-as-you-type
// (not a plain closed dropdown) so finding one title among many doesn't
// mean scanning a long list by eye - type a few letters, matches filter
// down, click the one you want. Alphabetical order, not grouped/ranked -
// an earlier version sorted "best authority-tier match" first (Execute ->
// analyst, Approve -> senior, etc.), but ties within a tier fell back to
// whatever order the underlying employee data happened to list them in,
// which read as arbitrary/scattered rather than useful. Alphabetical is
// the one order every user can predict without knowing the data.
const EmployeePicker: React.FC<IEmployeePickerProps> = ({ employees, value, onChange }) => {
  const options: IComboBoxOption[] = React.useMemo(() => {
    const distinctTitles = Array.from(new Set(employees.map(e => e.jobTitle)));
    return distinctTitles.sort((a, b) => a.localeCompare(b)).map(jobTitle => ({ key: jobTitle, text: jobTitle }));
  }, [employees]);

  return (
    <ComboBox
      label="Responsible"
      placeholder="Search job titles..."
      selectedKey={value}
      options={options}
      autoComplete="on"
      onChange={(_e, option) => {
        if (option) onChange(String(option.key));
      }}
    />
  );
};

export default EmployeePicker;
