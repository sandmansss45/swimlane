// Region values as they should appear in the region filter (All / South
// Africa / UK / US). Kept as plain strings rather than a closed union
// since the real list's actual values aren't confirmed yet - narrowing
// this later is a one-line change once we know.
export type Region = string;

export interface IEmployee {
  id: string;
  name: string;
  jobTitle: string;
  region?: Region;
}

// Lanes are always job titles, never a person's name - confirmed design
// rule. This formatter is only for the employee *picker*, which shows
// "Title — Name" so a human can tell people with the same title apart.
export function formatEmployeeLabel(employee: IEmployee): string {
  return `${employee.jobTitle} — ${employee.name}`;
}
