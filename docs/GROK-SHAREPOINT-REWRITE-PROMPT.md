# GROK BOT PROMPT — Rebuild VU Student Attendance Reporting on SharePoint

**Copy everything below the line into Grok Bot as the project brief.**

---

## MISSION

Rebuild the existing Windows desktop application **VU Student Attendance Reporting** (currently v3.2, VB.NET WinForms, SQL Server) so it does the **same jobs for teachers**, but stores and reads all operational data from **Microsoft SharePoint** instead of SQL Server.

You must recreate the whole app (preferred stack: modern .NET WinForms or WPF with Microsoft Graph / SharePoint REST, or CSOM if justified). Preserve behaviour, email workflows, Exemplar API integration, Outlook drafting, CSV/Excel imports, and OTA update checking.

There are **two app roles**:

1. **Owner / Manager** — connects to a SharePoint site and **creates / repairs the list structure, columns, views, and seed data** automatically so the app can run. Also performs admin data loads (agreements Excel, units CSV, teachers, templates, email routing, signature, housekeeping).
2. **Teacher** — connects to the already-provisioned SharePoint site and uses the app exactly like today’s teacher workflow: select class/student, send notices, log attendance, check units, profiling, etc.

Do **not** leave SQL Server as the primary store. SharePoint lists (plus a document library for the signature image / LEA PDFs if needed) are the system of record.

---

## PRODUCT CONTEXT

- Used by Victoria University Electrical / Engineering teaching & admin staff.
- Combines: student roster lookup, attendance intervention emails via Outlook, unit completion tracking, Exemplar profiling status, admin roster/units imports.
- Today: `ElectrotechnologyReports` SQL database on `DEVSQLCENTRAL.AD.VU.EDU.AU` via Windows Integrated Security (`My.Settings.SQLConString`).
- SharePoint NuGet packages exist in the old project but are **unused** for data — ignore them; implement clean Graph or modern auth.
- Outlook must remain installed and used to **Display()** drafts (user sends manually). Do not silently send unless explicitly designed later.
- Exemplar Profiling is a **separate HTTP API** (`https://api.profiling.exemplarsystems.com.au`) — keep it; only the local student/config data moves to SharePoint.

---

## ROLE MODEL (REQUIRED)

### A) Owner / Manager mode

**Purpose:** First-time setup and ongoing administration.

**On connect (site URL + Entra auth):**
1. Detect whether required lists/libraries exist.
2. If missing → create them with correct columns, types, indexes/views, and required seed rows.
3. If present but incomplete → add missing columns/lists (idempotent repair).
4. Seed default email templates (subjects listed below) if empty; seed `EmailSettings` rows for `Admin`, `Apptrain`, `Trades`, `Frank Offer` if missing; seed `Updates` singleton; seed `UEE30820units` catalogue if empty.
5. Show a clear “SharePoint ready / needs repair” status.

**Owner capabilities (map from current SettingsForm + Admin):**
- Edit EmailSettings (Admin / Apptrain / Trades addresses).
- Upload email signature image (store in a Document Library; list holds link or use library file `EmailSignature.jpg`).
- CRUD EmailTemplates (subject, body HTML with placeholders, help text).
- CRUD TeacherList.
- Upload **Agreements** Excel (full roster replace).
- Upload **Student Units** CSV (filter → staging → pivot into student unit flags + addresses).
- Set/update Student Database Date and Units Database Date.
- Reset yearly attendance counters.
- Delete orphan StudentLogs not present in Agreements.
- Toggle Mass Blockgroup Emailing.
- Manual AppTrain data reset (>14 days) equivalent.
- Optionally manage unit catalogue `UEE30820units`.

**Access:** Gate Owner mode with a strong local or Entra-group check (replace today’s password `Wpower84` / `Admin123` with proper roles). Do not hard-code those passwords in the new app; use Entra security groups or an Owner allow-list list in SharePoint.

### B) Teacher mode

**Purpose:** Day-to-day operations (map from Main + StudentUnits + SendOutlookEmail).

Teachers:
- Connect to the same SharePoint site (read/write on operational lists per permissions).
- Select Block Group → Student (or search Student ID).
- View student/employer/profile/attendance/unit alert/resit/profiling/employment.
- Choose email subject template → fill fields → Submit → Outlook draft + SharePoint side-effects (counters, dates).
- Student Investigation flow.
- BlockGroup mass email (if enabled).
- Open Student Units form (unit checkboxes, LEA letter, profiling correspondence).
- Amend Exemplar profiling email override.
- Employment status checkboxes.
- Check for app updates (GitHub Releases).
- Reconnect / show connection status.

Teachers must **not** see Owner provisioning/import tools unless also Owner.

---

## SHAREPOINT DATA MODEL (MAP FROM SQL)

Create these lists (names may be prefixed, e.g. `SAR_StudentLogs`, but keep stable internal keys). Column types should match usage.

### 1) `AgreementsDetails` (roster — Teacher read; Owner replace)

**Source today:** Excel upload truncates table then `SqlBulkCopy`.

**Columns (all used):**  
`Agreement ID`, `Student ID` (key for app), `Student Given Name`, `Student Family Name`, `Epsilon Start Date`, `Epsilon End Date`, `Student Personal Email`, `Student Personal Mobile`, `Employer Surname`, `Employer Given Name`, `Employer Contact Phone`, `Employer Email`, `Agreement Category`, `Course`, `Course Title`, `Course Status`, `Course Location`, `Block Group Code`, `Agreement Status`, `Agreement Task`, `Employer Name`, `Employer ABN`, `School Name`, `School Email`, `Org Unit`, `Agreement Name`, `Agreement Type Code`, `Training Plan Generated`, `Signed Training Plan Uploaded`, `Units for Employer Sign off`, `Progress Report Generated`, `Signed Progress Report Uploaded`, `All Units Resulted?`, `All Units Verified?`, `Number Units Completed`, `Course Hours Completed`, `Any Sanctions`, `Completion Training Plan Generated`, `Signed Completion Training Plan Uploaded`, `Department Email`, `Student VU Email`, `Actual Start Date`, `Actual End Date`, `Age of Agreement`, `Apprenticeship Client ID`

**Behaviours to preserve:**
- Distinct Block Group Codes populate class dropdown.
- Students listed as `"StudentID - Given Family"` for a block.
- Search by Student ID (normalize Excel-style `4608758.0` → `4608758`; strip BOM).
- Mass email pulls student (± employer) emails for a block.
- After agreements upload, update `Updates.StudentDatabaseDate`.
- Old app injects a test Frank Offer student if missing — replicate only if still required, behind a flag.

**SharePoint note:** Full replace = delete items in batches or use a staging list + swap; respect list view threshold / throttling. Prefer Graph batch deletes + creates. Index `Student ID` and `Block Group Code`.

### 2) `StudentLogs` (per-student operational state — Teacher R/W; Owner admin)

**Key:** `Student ID`

**Attendance:** `Absent`, `Late Arrival`, `Early Departure`, `Yearly_Absent`, `Yearly_Late_Arrival`, `Yearly_Early_Departure` (integers; increments on notices; reset to 0 on Term/Yearly report send)

**Reporting / AppTrain:** `LastStudentReportDate`, `Apptrain-Studentbeenemailed`, `AppTrain-Haveyourangstudents`, `AppTrain-OtherFormofcontact`, `AppTrain-OtherText`, `StudentAddress`

**Unit completion flags (booleans / Yes-No):**  
`UEECO0023`, `UEECD0007`, `UEECD0019`, `UEECD0020`, `UEECD0051`, `UEECD0046`, `UEECD0044`, `UEEEL0021`, `UEEEL0019`, `UEERE0001`, `UEEEL0023`, `UEEEL0020`, `UEEEL0025`, `UEEEL0024`, `UEEEL0008`, `UEEEL0009`, `UEEEL0010`, `UEEDV0005`, `UEEDV0008`, `UEEEL0003`, `UEEEL0018`, `UEEEL0005`, `UEECD0016`, `UEEEL0047`, `HLTAID009`, `UETDRRF004`, `UEEEL0012`, `UEEEL0014`, `UEEEL0039`

**Recommended redesign (still must behave the same in UI):**  
Optionally store units in child list `StudentUnitCompletions` (`Student ID`, `Unit_Code`, `Completed`) and project into checkboxes — OR keep wide columns if Owner provisioner creates all 29 Yes/No fields. Prefer child list for SharePoint limits, but Teacher UI must still show 29 checkboxes in order.

### 3) `StudentUnitsDatabase` (staging — Owner only)

`Student ID`, `units`  
Used only during CSV import: truncate/reload, then pivot into StudentLogs unit flags. Can be a temporary staging list cleared after pivot.

### 4) `TeacherList`

`Teacher_Full_Name`, `E_Number` (key), `Email`, `Contact_Number`, `Department`, `Highest_Certificate_Taught`, `Position`  
Main teacher dropdown filters Cert III teachers today — preserve that filter unless Owner configures otherwise.

### 5) `EmailSettings`

`SendTo` (unique: `Admin`, `Apptrain`, `Trades`, `Frank Offer`), `EmailAddress`, and signature via **Document Library** (not binary list column).

### 6) `EmailTemplates`

`EmailSubject` (unique), `EmailBody` (HTML with placeholders), `EmailHelp`

### 7) `Updates` (singleton)

`ID=1`, `DatabaseUpdateDate` (units CSV date), `StudentDatabaseDate` (agreements date), `VersionNumber` (legacy OTA fallback; GitHub Releases is primary)

### 8) `StudentEmploymentStatus`

`StudentID`, `EmploymentStatus` (e.g. unemployed / incorrectDetails), `DateofUnemployment`

### 9) `UEE30820units`

`Unit_Code`, `Unit_Title`, `ExceptionReport` (text used in profiling correspondence)

### 10) `ElectricalResit`

`Student ID`, `Unit`, `Resit date` — shown on Main as resit label

### 11) `ExemplarProfilingStudentDB`

`StudentID`, `ExemplarProfilingEmail` — override email for Exemplar API lookup (MERGE semantics)

### Libraries

- `SAR_Assets` or similar: `EmailSignature.jpg`, optional LEA outputs (today LEA PDFs go to a P: drive path — make configurable SharePoint folder or keep local path setting).

---

## CRITICAL SHARED HELPERS

### Student ID normalization (MUST KEEP)

Excel/CSV often emits `4608758.0`. Normalize before every lookup/write:
- Trim; strip BOM
- If parses as double → truncate to integer string
- Else use trimmed string

### No SQL JOINs today

All linking is by Student ID / app logic / EXISTS pivots. In SharePoint, implement the same in code (Graph queries + in-memory), not CAML joins.

---

## TEACHER MAIN SCREEN — BEHAVIOUR SPEC

### Startup sequence

1. Show loading UI.
2. Authenticate to SharePoint; verify lists ready (else tell Teacher to contact Owner / open Owner setup).
3. Require Outlook installed; warn about VPN / network as appropriate.
4. Load: Block Groups, Teachers, Units catalogue, Email Subjects, reset stale AppTrain data (>14 days clears last report date + AppTrain fields), show Student Database Date label, Exemplar API status, version label.
5. Hide Submit / Investigation until subject selected.
6. Optional silent update check (GitHub Releases).

### Controls and processes

| UI | Behaviour |
|----|-----------|
| Block Group combo | Load distinct Block Group Codes from AgreementsDetails. On change: clear student UI; load students for group; update mass-email visibility. |
| Student combo | On select: load AgreementsDetails profile fields; StudentLogs address + attendance counters; unit completion alert via CompletionChecker rules; ElectricalResit label; Exemplar profiling refresh; employment status checkboxes. |
| Search Student ID + Search button | LIKE/contains search on AgreementsDetails Student ID; select matching block + student. Enter key triggers search. |
| Email Subject combo | From EmailTemplates subjects. Drives `DisplayMode`: show/hide fields and Submit vs Investigation button. |
| Teacher combo | From TeacherList; fills teacher name/email labels. |
| Attendance/Punctuality, Classroom Engagement, Course Progression | Satisfactory / Not Satisfactory — used in templates. |
| Unit combo + Passed/Failed | From UEE30820units; unit report template. |
| Date + Time + Notes | Occurrence date/time/notes for templates. |
| **Submit** | Validate required fields per subject (see validation matrix). For Term/Yearly if no LastStudentReportDate → ask for last report date first. Then draft Outlook via template engine + apply StudentLogs side-effects + refresh counters + clear form fields. |
| **Student Investigation** | If subject is Investigation: prompt emailed? rang? other contact? + text → write AppTrain fields on StudentLogs; then Outlook draft (To Apptrain, CC Admin). |
| **BlockGroup Email** | If MassEmail enabled: ask include employers?; BCC all class (± employers); Outlook draft; signature + version. |
| **Check Students Units** | Open Student Units form for selected student. |
| **Amend/Assign Exemplar profiling email** | Lookup API; dialog; save override to ExemplarProfilingStudentDB; refresh. |
| **Student Re-Allocation Request** | Open Microsoft Form URL; optional Outlook template to student+employer CC trades/admin. |
| Employment checkboxes | Not employed → date + App-Train notified?; save StudentEmploymentStatus; may draft App-Train unemployment email; hide Submit when checked. Incorrect details → save status; hide Submit. |
| Check for Updates | GitHub latest release vs `v{Major}.{Minor}`; honor `min_required_tag` / force flags; download installer asset; exit after launch. |
| Reconnect | Re-auth / re-open Graph client; clear Exemplar token cache; refresh student if selected. |
| Exit | Close cleanly. |
| Owner / Admin entry | Role-gated entry to Owner console (replace UL password button). |
| Help `?` | Show EmailSubjectHelp from EmailTemplates.EmailHelp. |
| Issue Report | Open feedback URL (current Google Form). |

### Email subject list (must support)

1. Student Term Progress Report  
2. 2 Week Intention Letter  
3. 4 Week Intention Letter  
4. Course Withdraw Notice  
5. Student Behaviour Notice  
6. Overdue Fees - Warning  
7. Overdue Fees - Sanction  
8. Unit Withdraw Notice  
9. Absent Notice  
10. Late Arrival Notice  
11. Early Departure Notice  
12. Sent Back to Work Notice  
13. Student Unit Report  
14. Student Investigation *(uses Investigation button, not Submit)*  
15. Class Commencement Reminder  
16. Yearly Student Report  
17. Exemplar Profiling Outstanding Alert  

### Submit validation (required fields by subject)

Mirror current Main validation (missing fields are collected; preserve current UX — it builds `missingFields` messaging). Minimum required patterns:

- **Term Progress / Yearly:** subject, teacher, date, Att/Punct, Engagement, Progression (+ last report date picker if none stored)
- **2/4 Week Intention:** subject, teacher, date
- **Course Withdraw / Overdue Warning / Sanction:** subject, teacher
- **Behaviour / Unit Withdraw / Absent:** subject, teacher, date (+ unit for Unit Withdraw)
- **Late / Early / Sent Back To Work:** subject, teacher, date, time
- **Student Unit Report:** subject, teacher, date, unit, competent status
- **Investigation:** handled on Investigation button with contact prompts
- Others: follow existing DisplayMode field visibility

### Outlook routing

- Default notices: **To** employer email, **CC** student email, **BCC** Admin (from EmailSettings).
- Student Investigation: **To** Apptrain, **CC** Admin.
- Append signature image + version text to HTML body.
- Exemplar Profiling Outstanding Alert: append HTML summary of missing / not verified / employer verified / last card from Main labels.

### Template placeholders (replace in EmailBody)

`[employerFirstname]`, `[employerSurname]`, `[employerBusinessName]`, `[studentFirstname]`, `[studentSurname]`, `[studentID]`, `[LastStudentReportDate]`, `[Todaydate]`, `[AttPun]`, `[ClassRoomEngagement]`, `[CourseProg]`, `[AbsentLog]`, `[LateLog]`, `[EarlyLog]`, `[YearlyAbsentLog]`, `[YearlyLateLog]`, `[YearlyEarlyLog]`, `[Notes]`, `[TeacherName]`, `[TeacherEmail]`, `[ReportingTime]`, `[ReportingDate]`, `[UnitCode]`, `[UnitTitle]`, `[EmailedStudent]`, `[RangStudent]`, `[OtherContact]`, `[OtherContactText]`, `[TotalSchoolDays]` (prompt if used)

### StudentLogs side-effects on send (CRITICAL)

| Template | SharePoint write |
|----------|------------------|
| Student Term Progress Report | Set Absent, Late Arrival, Early Departure = 0; set LastStudentReportDate = today |
| Yearly Student Report | Same as Term |
| Absent Notice | Absent += 1; Yearly_Absent += 1 (create row if missing) |
| Late Arrival Notice | Late Arrival += 1; Yearly_Late_Arrival += 1 |
| Early Departure Notice | Early Departure += 1; Yearly_Early_Departure += 1 |
| Student Investigation | Read/write AppTrain contact fields before/during send |

Use read-modify-write with ETag/If-Match where possible to reduce lost updates.

### AppTrain stale reset

On startup: for StudentLogs where LastStudentReportDate older than 14 days, clear LastStudentReportDate and AppTrain contact fields (current `ResetApptrainData` behaviour).

### Profiling UI

- Resolve student via Exemplar API using override email if present else personal email / name.
- Show card summary; alert when missing cards ≥ 8 (“EXEMPLAR PROFILING ALERT!”).
- Keep Java login JAR fallback for token if still needed; prefer settings/env token.

---

## STUDENT UNITS FORM — BEHAVIOUR SPEC

Opened from Main for selected student.

- Show 29 unit checkboxes bound to StudentLogs flags (or child list).
- Changing a checkbox: require Owner/elevated confirmation (today password); then update SharePoint; re-run sequential prerequisite check.
- **Completion rules:** fixed unit order (`correctOrder` in CompletionChecker). All true → “UEE30820 - Completed”. Gap → “Student Is missing pre-requisite units”. Main `UnitAlertLbl` syncs on close.
- **Refresh Profiling:** HTTP unit progression % via Exemplar API + override email.
- **Generate / Email Authority to sit (LEA):** fill `LEATemplate.pdf`; save to configured folder; email subject `Congratulations - LEA Authorisation Form`.
- **LEA gating:** all units checked + address + assessment checks; block if avg profiling &lt;85% (UEEEL0012) or &lt;100% (UEEEL0039) or UEEDV0005/0008 EXP/DEMO &lt;50%; then show “Email Profiling Correspondence” instead (`Profiling Outcome - Insufficient Cards`) using `UEE30820units.ExceptionReport` guidance.

---

## OWNER CONSOLE — BEHAVIOUR SPEC

### Provision SharePoint structure

Button/action: **Connect & Ensure Structure**
- Create lists/libraries + columns + indexed fields + seed data.
- Output checklist of created/repaired items.

### Email settings & signature

- Update Admin / Apptrain / Trades addresses (Frank Offer used for legacy bug mail).
- Upload signature image to library.

### Email Templates CRUD

Add / Update / Delete / Reset fields; subjects drive Teacher combo.

### Teachers CRUD

Add / Update / Remove by E_Number; populate editor from selection.

### Agreements Excel import (Pipeline B)

1. Accept `.xlsx` only (reject `.xls`).
2. Find header row (contains Student ID + Agreement ID + Epsilon Start Date within first ~30 rows).
3. Require full column set (46 fields as in AgreementsDetails).
4. Normalize types; soft-null bad mobiles; skip blank/duplicate Student IDs.
5. Confirm → **replace** AgreementsDetails contents (batched).
6. Set Updates.StudentDatabaseDate = now/today.
7. Show progress and failure report.

### Student Units CSV import (Pipeline A)

Required CSV columns:  
`Student ID`, `Study Package Code`, `Grade Code`, `Student Study Package Status`, `Contact Student Address Line 1`, `Contact Suburb/Town`, `Contact State`, `Contact Postcode`

Filter rules (`CsvStudentUnitsFilter`):
- Keep rows where Grade ∈ {CBC, PP, GC} OR Status ∈ {Credited, Passed, Exempt} OR (Enrolled + CBC)
- Dedupe student+unit by priority: Exempt > Passed > Credited > Enrolled+CBC > grade-only

Then:
1. Clear staging `StudentUnitsDatabase`
2. Insert filtered (`Student ID`, unit code from Study Package Code) with normalized IDs
3. Build address string → update matching `StudentLogs.StudentAddress` (create StudentLogs row if needed)
4. Set Updates.DatabaseUpdateDate = today
5. Pivot: for each of 29 unit codes, set StudentLogs flags from presence in staging; insert StudentLogs for staging IDs missing from logs
6. Run AppTrain reset equivalent

### Housekeeping

- Reset all Yearly_* counters to 0
- Delete StudentLogs whose Student ID not in AgreementsDetails
- Manual AppTrain reset
- Toggle MassEmail setting (app setting or SharePoint config list)
- Show/edit database date labels

---

## EXEMPLAR API (KEEP — NOT SHAREPOINT)

Base URL setting default: `https://api.profiling.exemplarsystems.com.au`

| Call | Endpoint |
|------|----------|
| Find student | `GET /api/v1/users?roles=STUDENT&email=` or name params |
| Card summary | `GET /api/v1/users/{id}/cards/summary` |
| Qualifications | `GET /api/v1/qualifications`, user quals, qual detail, unit cards |
| Status update | `PUT .../qualifications/{qualId}` with status ACTIVE/INACTIVE/COMPLETED/WITHDRAWN |

Auth: Bearer token from settings/secure store; optional Java JAR login fallback (`ExemplarLogin.jar`). Qualification IDs configurable (old app has hardcoded GUID fallbacks — make settings).

Only **email override** persists in SharePoint (`ExemplarProfilingStudentDB`).

---

## OTA UPDATES (KEEP)

Primary: GitHub Releases API for repo used by the product (currently `Wirepower/VU-Student-Reporting`).
- Current version tag `v{Major}.{Minor}` from assembly
- Parse release body for `min_required_tag`, `force_update`/`mandatory_update`, `asset_name`
- Prefer assets `.msi` → `.msixbundle` → `.msix` → `.exe` → `.zip`
- Mandatory → download/launch/exit; optional → prompt

Legacy P-drive / Updates.VersionNumber path can be dropped or optional.

---

## AUTH & CONFIG SETTINGS

Replace `SQLConString` with:

| Setting | Purpose |
|---------|---------|
| `SharePointSiteUrl` | Site containing SAR lists |
| `TenantId` / app registration or delegated Graph auth | Entra auth |
| `MassEmail` | bool |
| `ExemplarApiBaseUrl` | string |
| `ExemplarApiToken` | secret |
| `ExemplarQualificationId` | string |
| `LeaOutputPath` or SharePoint folder | LEA PDF output |
| `GitHubRepo` | updates |

Teacher and Owner both authenticate with Microsoft accounts. Owner actions require Owner Entra group or `SAR_Owners` list membership.

---

## ARCHITECTURE REQUIREMENTS FOR THE REWRITE

1. **Data access layer** — single interface (`IStudentRepository`, etc.) so UI never embeds raw Graph calls. Implement SharePoint behind it.
2. **Idempotent provisioner** — Owner “Ensure Structure” is safe to re-run.
3. **Batching & throttling** — handle Graph 429s; never assume SQL-speed bulk.
4. **NormalizeStudentId** — one shared helper used everywhere.
5. **Do not port dead code** as features: unused `Settings.vb` twin, orphaned `StudentAmendment`/`BugReport` forms unless product owner asks; CompletionChecker must use the shared connection/config (no hard-coded SQL server).
6. **Preserve Outlook Display() draft behaviour** and placeholder names exactly so existing template HTML still works after migration of template content.
7. **Migrate data** — provide Owner tool or script notes to export SQL → SharePoint once (CSV per list) including unit flags and templates.
8. **Two-mode shell** — login → detect role → Teacher Home vs Owner Console; shared connection status.

---

## ACCEPTANCE CRITERIA

- [ ] Owner can point at an empty SharePoint site and provision all lists/libraries/columns/seeds until status = Ready.
- [ ] Teacher can select block/student and see the same profile fields as today.
- [ ] Each email subject drafts Outlook with correct recipients, placeholders, signature.
- [ ] Absent/Late/Early increments persist on SharePoint and show on Main.
- [ ] Term/Yearly send resets term counters and sets last report date.
- [ ] Units CSV + Agreements Excel Owner imports produce correct StudentLogs flags and roster.
- [ ] Student Units checkboxes + prerequisite alert + LEA gating work.
- [ ] Exemplar profiling + email override work without SQL.
- [ ] Mass block email works when enabled.
- [ ] Employment status flows work.
- [ ] Update check still works against GitHub Releases.
- [ ] No dependency on SQL Server for runtime operations.

---

## IMPLEMENTATION ORDER (FOLLOW THIS)

1. Auth + SharePoint provisioner (Owner) creating empty schema + seeds.  
2. Read-only Teacher roster/search/profile from AgreementsDetails + StudentLogs.  
3. EmailTemplates + EmailSettings + Outlook send path + counter side-effects.  
4. TeacherList, units catalogue, Student Units form + completion rules.  
5. Owner import pipelines (Agreements Excel, Units CSV) with progress UI.  
6. Employment, resits, Exemplar overrides, mass email, updates.  
7. Data migration utility from SQL export.  
8. Polish: reconnect, loading states, throttling, role gating, remove SQL.

---

## REFERENCE: CURRENT APP FILE MAP (BEHAVIOUR SOURCES)

Use these as behaviour oracles while rewriting (do not require VB.NET for the new app):

| Area | Files |
|------|-------|
| Main teacher UI | `Main.vb`, `Main.Designer.vb` |
| Outlook + counters | `SendOutlookEmail.vb` |
| Owner settings | `SettingsForm.vb` (live); ignore `Settings.vb` as legacy duplicate |
| Limited admin CSV | `Admin.vb` |
| Units UI | `StudentUnits.vb`, `CompletionChecker.vb` |
| Imports | `ExcelValidationModule.vb`, `CsvStudentUnitsFilter.vb` |
| Dates | `DatabaseDate.vb`, `LastReportDatePicker.vb` |
| Exemplar | `ExemplarProfilingApi.vb`, `ExemplarEmailOverrides.vb` |
| Resits | `ResitModule.vb` |
| Updates | `UpdateModule.vb` |
| SQL gateway (replace) | `SQLCon.vb`, `SQLError.vb` |
| Templates editor | `EmailTemplates.vb`, `EmailSubjectHelp.vb` |

Repo: VU Student Attendance Reporting, current version **v3.2**.

---

## FINAL INSTRUCTION TO GROK

Rebuild the application end-to-end for SharePoint with Owner/Manager provisioning and Teacher operations as specified. Prefer clarity, idempotent SharePoint setup, and faithful workflow parity over carrying forward SQL-era code structure. When a SharePoint limitation blocks a SQL pattern (MERGE, TRUNCATE+bulk, EXISTS pivot), implement the **same user-visible result** with Graph/batch/app-side logic and document the approach in README.
