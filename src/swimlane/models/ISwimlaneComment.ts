// A lightweight feedback channel (confirmed 2026-08-19) - anyone in the
// organisation viewing a swimlane can leave a comment suggesting a
// change, flagging an issue, or whatever else, without needing edit
// rights or going through the lock/edit flow at all (commenting is never
// blocked by an active lock - see IProgressIdLock - if anything, "this
// needs to change" is most useful to say WHILE something is locked for
// review). Append-only, same reasoning as IProgressIdLock: a comment is
// never edited or deleted once posted, so the Improvements tab is a
// durable record of who raised what and when, not a live discussion
// thread someone could quietly edit later.
//
// Scoped to progressId + region together, same reasoning as
// IProgressIdLock - a comment about the UK version of a flow shouldn't
// get mixed in with the US/SA versions living on the same Progress ID.
export interface ISwimlaneComment {
  id: string; // SharePoint list item ID
  progressId: string; // e.g. '9.6.1'
  region: string; // '' = commented while viewing "All"
  author: string;
  comment: string;
  postedAt: string; // ISO date string
}
