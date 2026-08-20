// A user-created name for a Category ID the static APQC_CATEGORY_NAMES
// lookup (apqcHierarchy.ts) doesn't already cover - e.g. a genuinely new
// top-level Category 14 the business adds that isn't part of the
// confirmed reference sheet. Same idea as IProcessGroupLabel one level
// up, kept as its own list rather than reusing that one, matching how
// this app already gives every hierarchy level its own dedicated list
// (Process Group Labels, Process ID Labels) rather than one shared
// generic table.
export interface ICategoryLabel {
  id: string;
  categoryId: string; // e.g. '14'
  name: string; // e.g. 'Manage Supply Chain'
}
