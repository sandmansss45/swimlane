# Swimlane Studio

A React/TypeScript/Vite web app for Quantum Leap Energy that visualizes
business process flows as swimlane diagrams (lane = job title,
column = sequence step), sourced either from mock fixture data or from
real SharePoint lists via Microsoft Graph.

## Stack

- React 19 + TypeScript + Vite 8, deployed as a static site to GitHub
  Pages via GitHub Actions (`.github/workflows/deploy.yml`, auto-deploys
  on every push to `main`).
- Fluent UI v8 for all form controls/dialogs/buttons.
- MSAL.js (`@azure/msal-browser`, `@azure/msal-react`) for Entra ID
  sign-in; Microsoft Graph for all SharePoint list reads/writes.
- No backend of any kind - this is a pure static SPA. All data access
  goes straight from the browser to Microsoft Graph.
- No automated test suite exists. Significant features are verified
  live via one-off Playwright scripts written to a local `pw-check/`
  folder, run manually, then deleted - not a persisted test suite.
  `npm run lint` runs oxlint; there is no `npm test`.

## Commands

```
npm run dev      # local dev server (Vite, http://localhost:5173)
npm run build    # tsc -b && vite build - ALWAYS run this before considering a change done
npm run lint     # oxlint
npm run preview  # serve the production build locally
```

## Architecture essentials

- **The canvas is a structured grid, not a free-form diagram.** Every
  step gets its own column (a "timeline slot"); lanes are
  `responsibleJobTitle` values. This is deliberate - it's what makes
  automatic arrow routing, dependency-order validation, region
  splitting, and CSV round-tripping all work. See the top-of-file
  comment in `src/swimlane/components/SwimlaneCanvas.tsx`.
- **DependsOn tokens are row numbers, not IDs.** E.g. `"9.6.1.1-3"`
  means "row 3 of the original source data" (header = row 1), not a
  Process Step ID. See `src/swimlane/utils/dependencyResolution.ts` and
  `parseDependsOn` in `src/swimlane/models/IProcessStep.ts`.
- **APQC hierarchy**: Category (`"9"`) → Process Group (`"9.6"`) →
  Process ID / one swimlane (`"9.6.1"`) → Activity / step
  (`"9.6.1.1"`) - see `src/swimlane/utils/apqcHierarchy.ts`. Static
  category/group/process-ID names live in that file; anything not in
  the static table falls back to a user-created custom label (its own
  small SharePoint list) or a derived name.
- **Region splitting**: one Process ID can hold several genuinely
  separate swimlanes side by side (UK/US/SA/Global), via
  `IProcessStep.region` - see `FlowRegionTabs.tsx`. `KNOWN_FLOW_REGIONS`
  in `models/IProcessStep.ts` is the one place to add a new region.
- **Append-only audit patterns**: swimlane locks (`IProcessIdLock`)
  and swimlane comments (`ISwimlaneComment`) are never edited or
  deleted once created - a new record is always added instead, so a
  full history survives. Both are scoped to `processId + region`
  together, not just `processId`.
- **Two data services, one interface** (`IDataService`):
  `MockDataService` (in-memory fixture data, used via "Use mock data
  (no sign-in)" on the sign-in screen) and `GraphDataService` (real
  SharePoint via Graph). Every new data-backed feature needs both
  implemented, or it'll only work in one mode.
- **SharePoint list titles are hardcoded constants** at the top of
  `GraphDataService.ts` (`PROCESS_LIST_TITLE`, `EMPLOYEES_LIST_TITLE`,
  etc.) and resolved to real Graph list IDs by exact display-name match
  at runtime. If a real list gets renamed in SharePoint, the matching
  constant here has to be updated too, or every read/write against it
  breaks. See `README-HANDOVER.md` for the full current list.
- **`erasableSyntaxOnly` is enabled** (tsconfig) - TypeScript
  constructor parameter-property shorthand
  (`constructor(private x: string)`) is NOT allowed. Use an explicit
  field declaration plus a constructor body assignment instead.

## Conventions this codebase follows closely

- **Comments explain WHY, not WHAT.** Every non-obvious decision has a
  comment citing the reason (often a real bug it fixed, or an explicit
  user decision with a date). Match that style - don't add comments
  that just restate the code.
- **Never fabricate SharePoint schema.** If a column's real display
  name on the live list isn't confirmed, the code says so explicitly
  (`TODO-CONFIRM: ...` in a comment) rather than presenting a guess as
  fact. Once confirmed against the real site, update the comment to
  `CONFIRMED <date> - ...`. Same standard applies to any other
  business-critical fact (real list names, real column names, real
  process data) - flag confidence level, don't guess and move on.
- **Small, real PRs over speculative abstraction.** No feature flags,
  no unused extensibility, no fields "for later." If something's not
  needed yet, it isn't built yet.

## Before considering any change done

1. `npm run build` - must succeed with zero TypeScript errors.
2. Manually verify the actual feature works - ideally live in a
   browser (mock mode needs no credentials: click "Use mock data (no
   sign-in)" on the sign-in screen). A clean build proves the code
   compiles, not that the feature works.
3. Clean up any temporary verification scripts before finishing.

## Where the deeper context lives

`README-HANDOVER.md` in this repo is the full handover/operations
document - infrastructure, Entra/auth configuration, the complete
SharePoint data model, known issues, and open risks. Read it before
making infrastructure- or auth-level changes.
