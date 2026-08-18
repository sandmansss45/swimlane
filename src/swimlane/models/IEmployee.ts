// Department values (e.g. "3101 - Security and Safety", "6002 -
// Finance") - shown as its own column in the Employees directory tab.
// Kept as plain strings rather than a closed union since the real list's
// actual values aren't a fixed set.
// CONFIRMED 2026-08-17: the real "QLE Existing Organisation" list has no
// dedicated region/geography column at all - Department is the only
// grouping field available, by explicit user choice. This was originally
// modeled as "region" throughout the app, which read as actual geography
// and confused real users looking at real department codes - renamed to
// match what the data actually is. There used to also be a toolbar filter
// that narrowed the Responsible/job-title picker by department - removed
// 2026-08-18 at a real user's request, since it sat in the same spot as
// filters that visibly change the diagram (see FlowRegionTabs) while
// itself only affecting one dropdown several screens away, with no
// visible feedback. The field itself stays; only the filter is gone.
export type Department = string;

// Lanes are always job titles, never a person's name - confirmed design
// rule, and by explicit user choice this list's real employee names
// aren't pulled into the app at all (job title is the only thing that
// matters here). No `name` field as a result - see EmployeePicker and
// EmployeesList, which dedupe by job title instead of differentiating by
// person.
export interface IEmployee {
  id: string;
  jobTitle: string;
  department?: Department;
}
