# VU Student Attendance Reporting  
## Complete Instruction Manual (End User + Admin Guide)

> Version: Repository state as of this document commit  
> Applies to: `master` and `release` branches

---

## Table of Contents

1. [Purpose of this application](#purpose-of-this-application)  
2. [What the application can do (feature summary)](#what-the-application-can-do-feature-summary)  
3. [System requirements and prerequisites](#system-requirements-and-prerequisites)  
4. [Install and first launch](#install-and-first-launch)  
5. [Main screen user guide](#main-screen-user-guide)  
6. [Email workflows and template behavior](#email-workflows-and-template-behavior)  
7. [Student Units screen guide](#student-units-screen-guide)  
8. [Student Amendment workflow](#student-amendment-workflow)  
9. [Settings and admin tools](#settings-and-admin-tools)  
10. [Exemplar profiling integration guide](#exemplar-profiling-integration-guide)  
11. [Updates and release behavior](#updates-and-release-behavior)  
12. [Troubleshooting guide](#troubleshooting-guide)  
13. [Operational notes for IT/admins](#operational-notes-for-itadmins)

---

## Purpose of this application

VU Student Attendance Reporting is a Windows desktop application used by teaching and support staff to:

- manage attendance-related interventions
- send structured communication emails to students/employers
- monitor student unit progression
- handle investigation/amendment workflows
- track Exemplar profiling progress
- keep reporting and audit data synchronized with SQL-backed records

The design goal is to centralize student operations in one desktop workflow instead of splitting work across separate systems.

---

## What the application can do (feature summary)

### Core capabilities

1. **Student selection and context loading**
   - filter by block group/class
   - search by Student ID
   - load key student/employer details from SQL

2. **Attendance/intervention communications**
   - generate Outlook draft emails using SQL-managed templates
   - log relevant attendance counters and report dates in SQL
   - support standard warning/progress/report scenarios

3. **Student investigation process**
   - structured prompts for follow-up actions
   - saves outcomes back to database
   - generates investigation communication

4. **Unit progression and competency management**
   - view/update per-unit status
   - refresh Exemplar unit progression percentages
   - generate and email authority-to-sit documentation

5. **Exemplar profiling integration**
   - API-based student lookup and profiling card summaries
   - optional student-specific profiling email override
   - support for production/staging API modes

6. **Administrative operations**
   - manage email settings, templates, teacher data
   - CSV/Excel-driven SQL maintenance operations
   - SQL connection and database date administration

7. **Application updating**
   - checks GitHub Releases for newer versions
   - can enforce mandatory update policies
   - launches installer assets directly from release

---

## System requirements and prerequisites

### Required

- Windows machine (app is WinForms/.NET desktop)
- SQL connectivity to configured environment
- Microsoft Outlook installed (required for email workflows)
- Microsoft Excel installed (for Excel-integrated operations)

### Included by current GitHub release installer

The current GitHub release pipeline packages:

- self-contained app runtime payload
- application files and dependencies
- Exemplar login JAR assets
- bundled Java runtime (`jre\bin\java.exe`, Temurin 17) for Exemplar login compatibility

### Not bundled

- Outlook/Excel desktop installations
- organization-specific VPN/network access requirements

---

## Install and first launch

1. Download the latest release installer from GitHub Releases.  
2. Run `StudentAttendanceReporting-Setup.msi`.  
3. Launch using Start Menu or Desktop shortcut (created by installer).  
4. On first run:
   - verify SQL connection is available
   - verify Outlook is installed and opens
   - verify student data loads via class selection or student ID search

If SQL is unavailable, the app will display a SQL error workflow and allow connection updates (admin paths).

---

## Main screen user guide

### Typical daily workflow

1. Open app.
2. Select **BlockGroup/Class Name** (or use **Search Student ID**).
3. Select student from list.
4. Review loaded context (student/employer, attendance/profiling labels).
5. Choose required **Email Subject** template.
6. Complete visible fields.
7. Click **Submit** (or **Student Investigation** for that specific workflow).
8. Optionally open **Check Students Units** for unit/profiling actions.

### Main controls (practical meaning)

- **Check for Updates**: triggers update check and optional installer launch.
- **Search Student ID + Search**: direct student lookup path.
- **BlockGroup/Class Name** dropdown: class-driven student selection path.
- **Email Subject**: selects communication scenario and adjusts form fields.
- **Submit**: validates and builds Outlook draft based on selected template.
- **Student Investigation**: dedicated investigation flow (only for that template).
- **Check Students Units**: opens Student Units screen.
- **BlockGroup Email**: class-wide email utility.
- **Amend/Assign Exemplar profiling email**: saves/removes student-specific Exemplar email override.
- **Student Re-Allocation Request**: opens re-allocation request flow (link + email draft).
- **Issue Report / Feature Request**: opens external reporting form.

---

## Email workflows and template behavior

Email bodies/help content are managed from SQL templates. Users select template via **Email Subject**.

### Common template options

- Student Term Progress Report
- 2 Week Intention Letter
- 4 Week Intention Letter
- Course Withdraw Notice
- Student Behaviour Notice
- Overdue Fees - Warning
- Overdue Fees - Sanction
- Unit Withdraw Notice
- Absent Notice
- Late Arrival Notice
- Early Departure Notice
- Sent Back to Work Notice
- Student Unit Report
- Student Investigation
- Class Commencement Reminder
- Yearly Student Report
- Exemplar Profiling Outstanding Alert

> Notes:
> - Exact list displayed comes from SQL template table.
> - Some fields become visible/required only for specific subjects.

### Submit behavior

When **Submit** is used:

- app validates visible/required fields
- builds template body from SQL with variable replacement
- appends signature and contextual status content as needed
- opens an Outlook draft with recipients/subject/body

### Investigation behavior

For **Student Investigation**:

- use **Student Investigation** button (not standard Submit path)
- prompts user through defined contact/actions flow
- logs outcomes to SQL
- prepares corresponding email communication

---

## Student Units screen guide

Open via **Check Students Units**.

### What it is for

- monitor and manage student unit completion/progression
- check profiling percentages
- produce and send authority-to-sit documentation

### Key actions

1. **Refresh Profiling %**
   - pulls latest unit progression from Exemplar endpoints
   - updates unit-level experience/demonstration displays

2. **Unit checkboxes**
   - update SQL-backed completion state
   - protected by admin override prompt for direct modifications

3. **Generate Authority to sit letter**
   - produces LEA/assessment letter output

4. **Email Authority to sit letter**
   - opens prepared Outlook draft with generated letter attachment

5. **Email Profiling Correspondence**
   - creates email draft for insufficient card/profiling communication

---

## Student Amendment workflow

The Student Amendment form supports class assignment changes and related communications.

### Supported actions

- add student to blockgroup
- move student between blockgroups
- remove student from current blockgroup
- report incorrect employer details

### Inputs typically required

- current/proposed blockgroup
- current/proposed class teacher
- reason for change
- proposed effective/start date
- submitter identity

### Output

- structured Outlook draft(s) to relevant stakeholders

---

## Settings and admin tools

Admin access is controlled through the admin entry path on main screen.

### Settings capabilities

- manage notification email addresses (Admin/App-Train/Trades)
- upload or update signature image
- manage teacher list
- open Email Templates management
- run data import/update operations (admin variants)

### Email Templates management

- select existing template
- edit subject/body/help text
- add/delete/reset template entries

### Advanced/admin maintenance (SettingsForm/Admin)

- update SQL connection string and restart
- update database date metadata
- run student/unit maintenance tasks
- upload agreements and perform SQL housekeeping actions

---

## Exemplar profiling integration guide

### What is integrated

- student lookup in Exemplar
- card summary/status retrieval
- unit progression endpoint calls
- optional qualification status update support in code

### Runtime behavior

- uses configured API base URL and token flow
- login JAR is selected by environment (production/staging logic)
- app can store student-specific override email in SQL

### Important for users

- if profiling is unavailable, app reports reason in status area
- with bundled Java, most Java runtime issues should be removed on release builds

---

## Updates and release behavior

The app checks GitHub Releases (`releases/latest`), not branches.

### Update becomes visible to users only when:

1. a newer release tag is published (e.g., `v3.1`, `v3.2`)
2. release has installer assets attached (`.msi`/`.exe`)

### Policy controls (release notes metadata)

Optional keys can enforce update behavior:

- `min_required_tag`
- `force_update`
- `asset_name`

---

## Troubleshooting guide

### 1) SQL Connection Error

Symptoms:
- startup SQL error screen
- data not loading

Actions:
1. verify network/VPN access
2. verify SQL server reachability
3. update SQL connection string (admin path)
4. restart app

### 2) Outlook/email issues

Symptoms:
- email draft cannot open
- Outlook automation errors

Actions:
1. verify Outlook is installed and opens manually
2. ensure default profile is configured
3. retry Submit after Outlook is ready

### 3) Exemplar profiling shows “Not configured” or login failure

Actions:
1. verify Exemplar settings and environment pairing (prod vs staging)
2. verify correct login JAR selection
3. verify bundled Java exists (`jre\bin\java.exe`) in install folder
4. verify credentials/token configuration path

### 4) Update check not offering update

Actions:
1. verify release is actually published (not draft)
2. verify release tag is newer than installed major/minor
3. verify assets are attached

### 5) Unit/profiling values do not refresh

Actions:
1. click **Refresh Profiling %**
2. verify student has valid mapping data for requested unit set
3. verify Exemplar API connectivity and qualification configuration

---

## Operational notes for IT/admins

### Branch and release model

- `master`: development baseline
- `release`: branch that triggers automated build/release workflow

### Current release automation includes

- version tag management (`vX.Y`)
- self-contained app publish
- WiX MSI build
- shortcut/icon setup
- Java runtime bundling
- GitHub Release publish with installer assets

### Recommended governance

- maintain template quality in SQL `EmailTemplates`
- regularly validate email routing addresses in settings
- validate SQL schema/table availability after environment changes
- keep Exemplar production/staging config aligned with intended JAR/token flow

---

## Quick reference: where to find more detail

- Release automation: `docs/RELEASE-AUTOMATION.md`
- MSI packaging details: `docs/MSI-Installer-Guide.md`
- Exemplar setup and debugging: `EXEMPLAR_API_SETUP.md`
- OTA policy options: `OTA_RELEASE_POLICY.md`
- Branch process notes: `WORKFLOW.md`

---

### End of manual

