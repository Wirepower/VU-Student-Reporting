# VU Student Attendance Reporting

Desktop WinForms application for managing student attendance reporting, class communications, unit tracking, and Exemplar profiling for VU Electrical/Engineering operations.

## What this application does

The app is used by teaching/admin staff to work from a single student-centric screen and quickly:

- find students by block/class and student ID
- review attendance and engagement indicators
- send structured email notices and reports
- track unit completion status in SQL
- run student investigation/amendment workflows
- view and refresh Exemplar profiling progress
- escalate communications (single student or class-wide)

In practice, it combines attendance operations, communication templates, and progression tracking into one desktop workflow.

## Core capabilities

### 1) Student search, class filtering, and profile context
- Filters students by block group/class.
- Supports student ID search.
- Displays key profile and contact context on selection.
- Loads related student records from SQL-backed data sources.

### 2) Attendance and intervention workflow
- Logs/updates attendance-related events (for example absent/late/early patterns).
- Supports investigation and amendment forms for intervention records.
- Tracks report dates and supporting status fields used for follow-up.

### 3) Structured communication and email automation
- Builds template-based messages for common educator/admin scenarios.
- Opens Outlook compose windows with recipients, subject, and prefilled HTML body.
- Includes system signature imagery and version text in generated communications.
- Supports class-wide/mass email actions where required.

### 4) Unit completion and progress management
- Student Units form allows unit-level status review and updates.
- Persists unit completion state to SQL.
- Supports admin override gating for protected checkbox updates.
- Refreshes completion labels and related UI state after updates.

### 5) Exemplar profiling integration
- Integrates with Exemplar APIs for student/profile lookup and progression data.
- Retrieves card summaries and unit progression metrics.
- Surfaces profiling status in the UI and supports profiling refresh actions.
- Includes configuration for production/staging API behavior (see `EXEMPLAR_API_SETUP.md`).

### 6) Update delivery and field rollout
- App checks GitHub Releases (`/releases/latest`) for OTA update availability.
- Can enforce optional mandatory update policy via release metadata.
- Downloads and launches preferred installer assets (`.msi`/`.exe`) from releases.

## Typical user flow

1. Launch app and load SQL-backed data.
2. Select class/block and student.
3. Review attendance/profiling signals.
4. Send appropriate communication template(s).
5. Update unit/intervention records where needed.
6. Save/close with SQL state and labels refreshed.

## Technology and integration points

- **Platform:** .NET 8 WinForms (Windows desktop)
- **Data:** Microsoft SQL Server (via `Microsoft.Data.SqlClient`)
- **Email client integration:** Microsoft Outlook COM interop
- **Spreadsheet/PDF helpers:** Excel interop and PDF utilities
- **External profiling system:** Exemplar APIs (token-based)
- **Packaging:** WiX MSI + self-contained EXE release assets

## Branches

- `master`: main development branch (kept in sync with `release` in this repository).
- `release`: release-trigger branch used by GitHub Actions to build/publish installer releases.

## Update and release model

The application checks **GitHub Releases latest** (`/releases/latest`) for OTA updates.

- Branch pushes alone do not update field clients.
- A **published GitHub Release** with a newer `vX.Y` tag and attached installer assets is what clients consume.

## Automated release pipeline

When changes are pushed to `release`, the workflow at `.github/workflows/release.yml`:

1. Determines the next release tag in `vX.Y` format (or uses manually supplied version).
2. Publishes a self-contained Windows payload (`win-x64`).
3. Builds MSI with WiX.
4. Publishes a GitHub Release with:
   - `StudentAttendanceReporting.exe`
   - `StudentAttendanceReporting-Setup.msi`

## Installer behavior

The MSI creates:

- Desktop shortcut
- Start Menu shortcut

Both use `VU Support Hub_Desktop Icon-Favicon.ico`.

## Additional docs

- `docs/RELEASE-AUTOMATION.md` - release workflow details and branch behavior.
- `docs/MSI-Installer-Guide.md` - packaging/installer background and guidance.
- `EXEMPLAR_API_SETUP.md` - Exemplar API setup and environment details.
- `OTA_RELEASE_POLICY.md` - release metadata options for OTA policy.
