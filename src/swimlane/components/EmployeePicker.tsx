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
      // `text`, not `selectedKey` - a step's real responsibleJobTitle is
      // often a richer 3-part string ("Title / Code - Dept / Region", see
      // formatLaneLabel in SwimlaneCanvas.tsx) than anything this list of
      // options can offer, since IEmployee only ever carries a bare job
      // title (no region field exists in the real Employees list at all -
      // see IEmployee.ts). selectedKey only displays text when it exactly
      // matches an option's key, so it showed BLANK for every step with
      // one of these richer values - not just after a drag, on every
      // single edit-panel open, which is what actually made a drag look
      // like it "didn't change the job title" in the box. `text` instead
      // just displays whatever the step's current value literally is,
      // matching how the Action/Action type fields already work.
      //
      // allowFreeform is required for that same `text` prop to actually be
      // typeable - without it, Fluent treats `text` as fully authoritative
      // on every render, so each keystroke got redrawn straight back to
      // the old value before the user could see what they'd typed, and
      // the options list never got a chance to narrow (confirmed live:
      // typing "Chief" left the field showing the old value and all
      // options unfiltered). onChange below still only ever commits a
      // REAL option's key, never freeformValue, so a job title that isn't
      // in the Employees list still can't be saved - allowFreeform only
      // unlocks live typing/filtering, not arbitrary freeform values.
      text={value}
      options={options}
      autoComplete="on"
      allowFreeform
      onChange={(_e, option) => {
        if (option) onChange(String(option.key));
      }}
    />
  );
};

export default EmployeePicker;
