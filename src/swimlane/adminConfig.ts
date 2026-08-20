// A single shared password gating "Unlock swimlane" - a soft deterrent
// against casual/accidental unlocking, NOT real access control. This
// check runs entirely in the browser: anyone who opens DevTools and
// reads the bundled JS (or just watches someone else type it once) can
// see/learn this value, and nothing on the SharePoint/Graph side
// actually enforces it - the app has no server-side permission check at
// all yet. See the "Future architecture considerations" note in
// README-HANDOVER.md, which already flags that real access control needs
// enforcement beyond an in-app check, not just this.
//
// Change this to your own value before relying on it. Takes a rebuild +
// redeploy to take effect, same as any other compile-time constant here
// (e.g. ENTRA_CLIENT_ID in auth/authConfig.ts).
export const ADMIN_UNLOCK_PASSWORD = 'change-me';
