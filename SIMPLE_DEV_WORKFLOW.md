# Simple Dev Workflow (/dev URL)

Use this for normal development.

1. Open your Apps Script web app **/dev** URL in the browser.
2. Make code changes locally.
3. Push changes:

```bash
npm run dev:push
```

4. Hard refresh browser:

```text
Cmd+Shift+R
```

The `npm run dev:webapp` helper opens the current `@HEAD` deployment with the supported clasp command.

5. Test. Repeat.

## Helpful Commands

- Push + reminder message:

```bash
npm run dev:loop
```

- Open web app page from terminal:

```bash
npm run dev:webapp
```

## Release Commands (only when publishing /exec)

- List deployments:

```bash
npm run release:list
```

- Create a new version:

```bash
npm run release:version
```

Legacy release script is still available:

```bash
npm run release:deploy:legacy
```

## Read-only offline mission export

The authenticated GAS Tools panel has **Offline Download**. Choose 7 or 14 days,
prepare the package, review every mission's dependency report, then use the
download link to save JSON. The window includes today in the spreadsheet's time
zone and excludes the displayed upper-bound date. There is no 15-mission cap.

An active staff account is required. Pilots receive missions assigned to them
as pilot/copilot on every leg; existing supervisor/coordinator roles can export
other assignments. Only approved/already-started missions (including their
completed legs) are included, never pending or cancelled missions. Export does
not approve missions or assert current airworthiness/fitness to fly.

The collection has `kind: flightapp.offline-package`, `schemaVersion: 1`,
`missions[]`, shared `references`, a source revision, and per-mission
`completeness.missionResults`. A package is ready only when it is nonempty and
every included mission has its required data. Incomplete packages can be
downloaded for diagnosis and remain explicitly not ready.

Each mission includes the existing summary/detail contracts and a `legs` map.
Each leg contains its W&B state, normalized route, nullable flight-log baseline,
and shared-reference keys. Reconstruct W&B by combining
`references.wbConfigurations[leg.wbConfigurationRef]`, `leg.wb`, and
`envelopeData: references.envelopes[leg.envelopeRef]`. Performance setups point
to `references.performanceReferences.global` rather than duplicating calculation
tables. A null log baseline is valid before Briefing creates a flight-log row.
The source revision detects changed input data, not cryptographic authenticity.

Read-only chain:

- Export-specific Dispatch query, staff/assignment checks and schema checks.
- `_getPilotStartupDataReadOnly_` -> `_getDropdownData_(true)` ->
  `_mxReadDispatchMaintenanceStatus_(..., true)` ->
  `_mxReadScheduleData_(..., true)`. This path validates existing schema and
  never calls `_mxEnsureFrameworkSheets_`.
- Existing `getMissionById`, `initializeWB`, `getWbEnvelopeByAircraft`,
  `getFlightRouteData`, `getFlightLogData`, `getPerformanceSetup`, and
  `_toolsDutyReadConfigMap_` only read Sheets or compute values in this path.
- Normal public startup/dropdown/maintenance readers retain schema
  initialization. No save functions, HTTP allowlist additions, notifications,
  inventory changes, triggers, or file creation in Drive are used.

Local validation (does not contact GAS):

```powershell
node --test scripts\test-offline-package.js
git diff --check
```

Tests execute the actual backend readers against write-rejecting mock Sheets,
including multiple aircraft, all legs, more than 15 missions, missing references,
normal-reader compatibility, JSON safety and the download UI. Node tests are
excluded from clasp uploads.

After an explicitly approved future push, manually verify organizational sign-in,
assignment filtering, real sheet headers, both windows, dependency reports, and
JSON saving to Files on the issued iPad. No production push/deploy is necessary
for local tests. PWA import/storage and upload synchronization are not implemented
by this exporter. Remote charts/photos/maps and PWA assets are not embedded.

Predeployment checks also validate the fixed column positions still used by the
existing flight-log/fuel-log readers. Reordered columns are reported as incomplete,
never repaired. Runway dependency checks use the existing performance setup's
surface translator (for example, ASFALTO to Paved).

For development testing, after explicit approval:

```powershell
Set-Location 'C:\Users\Windows\OneDrive\Desktop\Flight_App'
git branch --show-current
git status --short
node --test scripts\test-offline-package.js
git diff --check
clasp status
npm run dev:push
npm run dev:webapp
```

Confirm the branch is `agents/offline-package-exporter` and review the complete
clasp upload list first: pushing uploads all clasp-selected local files, including
pre-existing API/Enroute changes and the local iframe test HTML, not just the
exporter. The `/dev` URL runs current saved code and requires script-editor access;
it does not update an existing versioned `/exec` deployment. Do not create a new
version or change domain access settings for this development check.

Measure both elapsed execution time and downloaded JSON size with real Sheets,
especially the 14-day window. Existing readers reread entire sheets per mission
and per leg; mock tests cannot establish real Apps Script latency or response-size
capacity. A timeout/transport failure must produce no ready/downloadable package.
Large airport/reference collections and remote asset URLs also require testing on
the issued iPad. Treat exported files as sensitive organizational data.
