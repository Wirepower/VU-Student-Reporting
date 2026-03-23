# Setup / MSI (Visual Studio Installer Project)

This folder contains **`Setup.vdproj`** — the Visual Studio **Installer Projects** setup that builds **`Setup.msi`**.

## Build output

- **Debug:** `Setup\Debug\Setup.msi`
- **Release:** `Setup\Release\Setup.msi`

Build folders are **gitignored** (standard Visual Studio); only the **`.vdproj` source** is in Git. Build the MSI locally or in CI when needed.

## Before building the Setup project

1. Publish the app (self-contained recommended):

   ```powershell
   powershell -ExecutionPolicy Bypass -File "..\build\Publish-ForInstaller.ps1" -SelfContained
   ```

2. In Visual Studio, **Rebuild** the **Setup** project (solution build includes it).

## Notes

- **Icon:** `VU Support Hub_Desktop Icon-Favicon.ico` at the repo root is referenced by the installer.
- **Published items** output group pulls **`PublishItems`** from the main VB project.
- Full steps: **`docs/MSI-Installer-Guide.md`**

## Old location

If you previously used **`..\Setup\Setup.vdproj`** (sibling folder under `repos\`), that copy is superseded by **`Setup\Setup.vdproj` inside this repository**.
