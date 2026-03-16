# OTA release policy (GitHub Releases)

This application now checks GitHub Releases for OTA updates.

## Current release tag

The app currently uses:

- `v2.0sql`

When preparing a new build, use an incremented tag like:

- `v2.1sql`
- `v3.0sql`

## Release notes metadata (optional but recommended)

The updater reads simple key/value lines from the release notes body.

Supported keys:

- `min_required_tag` (or `min-required-tag`)
- `force_update` (or `force-update`)
- `asset_name` (or `asset-name`)

Example release body:

```text
Release highlights:
- Added API student profiling
- Improved high-DPI layout

min_required_tag=v2.1sql
force_update=true
asset_name=StudentAttendanceReporting-Setup.exe
```

## Behavior

- If latest release tag is newer than current tag, app offers update.
- If `min_required_tag` is above current tag, update is mandatory.
- If `force_update=true` and latest tag is newer, update is mandatory.
- If `asset_name` is provided, that asset is preferred for download.

## Asset selection fallback

If `asset_name` is not provided, updater picks the first matching asset in this order:

1. `.msi`
2. `.msixbundle`
3. `.msix`
4. `.exe`
5. `.zip`
