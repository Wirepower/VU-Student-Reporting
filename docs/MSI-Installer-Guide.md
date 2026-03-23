# MSI installer for Student Attendance Reporting

This app targets **.NET 8 (Windows)** and uses **COM** (Outlook/Excel interop). An MSI can ship your **application files** and optionally the **.NET runtime** (self-contained publish). It **cannot** legally bundle Microsoft Office or install Outlook/Excel for the user.

## What “install everything” means here

| Component | In MSI? | Notes |
|-----------|---------|--------|
| Your app + NuGet DLLs + `jre\`, PDFs, JARs | Yes | Use publish output from `build/Publish-ForInstaller.ps1` |
| .NET 8 Desktop Runtime | Optional | Either bundle **self-contained** publish (larger) or add **.NET 8 Desktop Runtime x64** as a prerequisite/bootstrapper |
| Microsoft Outlook / Excel | **No** | Must already be installed on the PC for COM features |
| VPN / P: drive | **No** | Environment/network; document for IT |

## Recommended: Visual Studio Installer Projects → MSI

1. Install **Visual Studio 2022** (you have it).
2. Install extension: **“Microsoft Visual Studio Installer Projects”** (Microsoft) from Extensions → Manage Extensions.
3. Open `Student Attendance Reporting.sln`.
4. **Fix or add the Setup project**
   - If the solution references `..\Setup\Setup.vdproj` and that path does not exist, either:
     - Move your `Setup.vdproj` into `VU-Student-Reporting\Setup\` and edit the `.sln` line to:  
       `"Setup\Setup.vdproj"`  
     - Or: **Add → New Project → Setup Project**, name it `Setup`, place it under this repo’s `Setup\` folder.
5. **Build the payload** (choose one):

   **Framework-dependent** (smaller MSI; users need .NET 8 Desktop Runtime):

   ```powershell
   .\build\Publish-ForInstaller.ps1
   ```

   **Self-contained** (larger MSI; includes .NET runtime in the folder — fewer prerequisites):

   ```powershell
   .\build\Publish-ForInstaller.ps1 -SelfContained
   ```

   Output defaults to: `publish\installer-payload\win-x64\`

6. In the **Setup** project:
   - Right-click **Application Folder** → **Add → Project Output…** is for the *main* project’s Primary Output **or** use **Add → File…** / **Folder** and add **all files** from `publish\installer-payload\win-x64\` (self-contained is usually easiest as “whole folder”).
   - Add shortcuts (Start Menu / Desktop) to the main `.exe`.
7. **Prerequisites** (if you used framework-dependent publish):
   - Setup project **Properties** → **Prerequisites** → e.g. **.NET Desktop Runtime 8.x (x64)**  
   - Or ship self-contained and skip separate runtime prerequisite.
8. Build the Setup project → produces **MSI** (and optionally setup.exe bootstrapper).

## Why the old “Setup” entry showed “load failed”

The solution previously pointed to `..\Setup\Setup.vdproj` (outside this repo). The installer project must live at a path that exists, or remove the stale project from the solution.

## Signing and IT deployment

- For domain rollout, IT may require a **code-signed** MSI and a **silent** install switch — configure that in the Setup project or with your packaging tool.

## Alternatives

- **WiX Toolset**: More control, XML authoring; good for CI pipelines.
- **MSIX**: Modern packaging; different distribution model than classic MSI.

If you want a **WiX** project committed next to this repo, say so and we can add a minimal `wxs` that harvests the publish folder.
