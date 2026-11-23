# Common Ground

A personal, local-first product-management sprint board with a high-contrast dark theme. No build step; ExcelJS powers workbook export.

## Run

```sh
npm install
npm start
```

Open http://localhost:3000. Requires Node.js 20 or later. Set `PORT` to use another port.

## Planning

- Create quarterly sheets with an initial sprint date, mobile version, and sprint count. Fresh installations and all new sheets start with no initiatives. Reloads preserve your saved work; unchanged samples from older installations are removed.
- Every new quarterly sheet starts at Sprint 1 and counts forward without resetting at calendar boundaries. Saved sheets with duplicate numbers caused by the old calendar-reset behavior are repaired on load or backup import, without changing dates or initiatives.
- Sprints are 14 days, including their start and end dates. Wednesday starts are recommended. Six sprints cover 12 weeks; seven cover a full quarter.
- Mobile cutoff is the Monday strictly before the sprint starts. Release is seven days after that cutoff, the following Monday. For Wednesday sprints this is five days after the start, not after the sprint ends.
- Mobile minor versions increment once per sprint. Click a sprint header to override any date or version independently.
- Add, rename, or remove teams. Removing a team with work requires selecting a replacement team.
- Click an initiative to edit its responsible team, sprint, description, or status: planned, to do, in progress, done, dropped. Selecting team dependencies shows linked copies in those teams' rows in the same sprint, including in Excel. Editing any copy updates the same task; removing a dependency removes its linked copy. Dragging a linked copy moves the shared sprint without changing its responsible team. Overview totals and the Excel Initiatives tab count the task once. There are no people assignments or creator labels.
- Drag initiatives between cells, or change their sprint in the details dialog on touch devices or using a keyboard. Later-sprint moves ask whether to preserve a spillover marker. Markers link to the current initiative and follow subsequent moves.
- Search and filter by status or responsible team. The release calendar provides a compact schedule view.
- Click **Expand board** to give the board the entire viewport. Filters, editing, and drag-and-drop remain available. Click **Exit expanded view** or press Escape to return; Escape closes an open editing dialog first.
- Use the sidebar toggle at the top left of the page to collapse or expand navigation. Your choice is remembered in this browser.
- Changes persist in this browser's local storage. Export/import JSON backups from the footer or settings button. Import replaces all sheets after confirmation.
- Use **Delete sheet** in the board footer to permanently remove the current sheet after confirmation. Deleting your last sheet leaves no sheets, with options to create a new one or import a backup. This empty state persists after reload. Export a backup first if you need to keep the deleted plan.

## Excel Export

Click **Export Excel** in the board toolbar to download the active sheet as a real `.xlsx` workbook. It includes all teams, sprints, and initiatives, regardless of current filters. The browser bundle is served locally and loaded only when exporting.

Exported mobile versions include a `v` prefix (for example, `v7.162`) to avoid Excel's numeric-text warning. Each sprint's top six header rows have a distinct tint and matching dark header. All merged team labels use the same light-blue fill; initiative cells remain white and separator bars retain their uniform formatting.

The returned workbook exports the entire supplied sheet, independent of UI filters, with three tabs: **Board** (teams by sprint, frozen headers and team column), **Initiatives** (one detailed row per initiative or spillover snapshot), and **Sprint schedule**. Dates are Excel date cells; user content is literal text, never formulas. Spillover records retain their original placement and status and identify the current destination sprint and team.

The Board tab follows a planning-grid layout: one sprint header per column, five compact schedule rows, dark separators between teams, vertically merged team labels, and one initiative per cell. Team blocks contain at least three rows. Status, responsible team, description, and spillover details are included in cell notes and the Initiatives tab. Calendar-quarter labels are used when a sheet spans quarters; fiscal-quarter labels are not inferred.

Excel exports are for sharing and reporting, not backup import; keep using JSON for restorable backups. The reusable `excel.js` builder accepts an ExcelJS Workbook constructor for both browser and Node use.

## Scope

This is an individual-use planning app. No login, profiles, cloud sync, or collaboration features are included. Old backups remain loadable; any legacy owner/stakeholder fields are retained in JSON for data preservation but are neither shown nor included in Excel exports. Do not rely on browser storage as the only copy of critical plans. Optional Google Fonts fall back to system sans-serif fonts when offline.

## Tests

```sh
npm test
```

Tests cover schedule calculations, version increments, spillover moves, backup validation, and serialized/reloaded Excel exports.

## GitHub Pages

`npm run build` creates a static `dist/` site including the ExcelJS browser bundle. The Pages workflow tests, builds, and deploys pushes to `main`. In repository Settings > Pages, choose GitHub Actions as the source. Private repositories require a GitHub plan that supports Pages; repository privacy alone does not make the Pages website private.

Localhost and the hosted site have separate browser storage. Export a JSON backup locally and import it on the hosted site to transfer your plans. No browser-saved plans are included in the deployment.
