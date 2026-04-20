# Release automation (release branch -> GitHub Release)

This repository supports automated release publishing from the `release` branch.

## What triggers a release

- Any push to the `release` branch runs `.github/workflows/release.yml`.
- You can also run it manually from **Actions** using `workflow_dispatch`.

## What the workflow does

1. Determines the next tag in `vX.Y` format.
   - If no tags exist, starts at `v1.0`.
   - Otherwise increments minor (`v3.0` -> `v3.1`).
   - Manual runs may provide an explicit `vX.Y` version.
2. Publishes a self-contained `win-x64` application payload (`.exe`) using `build/Publish-ForInstaller.ps1`.
3. Builds `StudentAttendanceReporting-Setup.msi` via WiX using `build/Build-WixMsi.ps1`.
4. Creates and pushes the tag.
5. Publishes or updates the GitHub Release with both assets:
   - `StudentAttendanceReporting.exe`
   - `StudentAttendanceReporting-Setup.msi`

## Why this matches updater behavior

The application checks **GitHub Releases latest** (`/releases/latest`), not branches.

That means:
- Pushing to `release` branch alone does not update clients.
- A **published release** with a newer `vX.Y` tag is what clients see.

## WiX requirement

The workflow installs WiX Toolset automatically on the GitHub runner.
No Visual Studio installer extension is required for CI releases.
