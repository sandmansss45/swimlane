// Region values as they should appear in the region filter (All / South
// Africa / UK / US). Kept as plain strings rather than a closed union
// since the real list's actual values aren't confirmed yet - narrowing
// this later is a one-line change once we know.
// CONFIRMED 2026-08-17: the real "QLE Existing Organisation" list has no
// dedicated region column - this is sourced from Department instead (e.g.
// "3101 - Security and Safety"), by explicit user choice. Kept the field
// name "region" rather than renaming it throughout the app, since nothing
// user-facing ever renders the literal word "region" - RegionFilter just
// shows "All" plus whatever distinct values are actually present.
export type Region = string;

// Lanes are always job titles, never a person's name - confirmed design
// rule, and by explicit user choice this list's real employee names
// aren't pulled into the app at all (job title is the only thing that
// matters here). No `name` field as a result - see EmployeePicker and
// EmployeesList, which dedupe by job title instead of differentiating by
// person.
export interface IEmployee {
  id: string;
  jobTitle: string;
  region?: Region;
}
