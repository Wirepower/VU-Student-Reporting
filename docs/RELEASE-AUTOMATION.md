# Release automation (release branch -> GitHub Release)

This repository supports automated release publishing from the `release` branch.

## What triggers a release

- Pushes to the `release` branch run `.github/workflows/release.yml` for application/release content changes.
- Markdown/docs-only updates and direct edits to the release workflow file are ignored on push (to avoid accidental release publishes).
- You can also run it manually from **Actions** using `workflow_dispatch`.

## Manual release prompts (workflow_dispatch)

When running manually from Actions, the workflow now asks for:

- `version` (optional `vX.Y`; if blank, next minor is auto-selected)
- `force_update` (`true` or `false`)
- `min_required_tag` (optional `vX.Y`)
- `asset_name` (preferred installer file name for in-app updater)
- `release_notes` (optional plain-text highlights)

The workflow writes these values into the GitHub Release body as OTA metadata:

```text
# force_update: true = mandatory update when this release is newer than installed version.
force_update=true

# min_required_tag: clients below this version must update before continuing.
min_required_tag=v3.1

# asset_name: preferred installer asset selected by in-app updater.
asset_name=StudentAttendanceReporting-Setup.msi
```

## What the workflow does

1. Determines the next tag in `vX.Y` format.
   - If no tags exist, starts at `v1.0`.
   - Otherwise increments minor (`v3.1` -> `v3.2`).
   - Manual runs may provide an explicit `vX.Y` version.
2. Publishes a self-contained `win-x64` application payload (`.exe`) using `build/Publish-ForInstaller.ps1`.
3. Builds `StudentAttendanceReporting-Setup.msi` via WiX using `build/Build-WixMsi.ps1`.
4. Creates and pushes the tag.
5. Publishes or updates the GitHub Release with both assets and OTA metadata:
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
