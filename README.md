# VU Student Attendance Reporting

Desktop WinForms application for student attendance reporting, communications, and Exemplar integration.

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

## Installer shortcuts

The MSI creates:

- Desktop shortcut
- Start Menu shortcut

Both use `VU Support Hub_Desktop Icon-Favicon.ico`.

## Additional docs

- `docs/RELEASE-AUTOMATION.md` - release workflow details and branch behavior.
- `docs/MSI-Installer-Guide.md` - packaging/installer background and guidance.
