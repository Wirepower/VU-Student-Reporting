# Exemplar API integration notes

The app now includes read-only student profiling lookup against Exemplar APIs.

## Staging vs production

- **ExemplarLogin.jar** = **production** login JAR. Use with the default API URL.
- **ExemplarLoginDev.jar** = **staging/development** login JAR. Use when `EXEMPLAR_API_BASE_URL` is set to the staging API URL.

The app picks the JAR automatically: if `EXEMPLAR_API_BASE_URL` is set, it uses **ExemplarLoginDev.jar** (embedded or in the app folder); otherwise it uses **ExemplarLogin.jar**. You can embed one or both in the project; the correct one is chosen at runtime.

- **Default API URL** (production): `https://api.profiling.exemplarsystems.com.au` — use **ExemplarLogin.jar** with it.
- **Staging**: set `EXEMPLAR_API_BASE_URL` to the staging API URL and use **ExemplarLoginDev.jar**. A staging token returns 401 against the production URL.

If you see "401 after token refresh", the JAR and API do not match: use ExemplarLogin.jar for production (default URL) or set staging URL and use ExemplarLoginDev.jar.

## Setup (one-time: you do this; users do nothing)

**1. Hardcode credentials**  
In `ExemplarProfilingApi.vb`, set the two constants:

- `ExemplarApiUsername` = your Exemplar login email  
- `ExemplarApiPassword` = your Exemplar password  

That is the only place you need to set credentials. The app uses these to get a Bearer token automatically. (Env vars `EXEMPLAR_API_USERNAME` / `EXEMPLAR_API_PASSWORD` can override these if set on a machine, e.g. by IT.)

**2. Embed the login JAR (so users don’t need the JAR)**  
- **Production**: save the production JAR as **ExemplarLogin.jar** in the project folder.  
- **Staging**: save the staging/development JAR as **ExemplarLoginDev.jar** in the project folder. You can embed one or both; the app uses ExemplarLoginDev.jar when `EXEMPLAR_API_BASE_URL` is set, otherwise ExemplarLogin.jar.  
- The `.vbproj` already embeds `ExemplarLogin.jar` and `ExemplarLoginDev.jar` when those files exist. Rebuild and the JAR(s) are embedded. Users get one .exe (and Java on the machine); no JAR setup for them.

**3. Optional environment overrides (for deployment)**  
- `EXEMPLAR_API_BASE_URL` = override API base URL if needed.  
- `EXEMPLAR_LOGIN_JAR_PATH` = path to JAR if you don’t embed it (app uses this or a JAR in the app directory).  
- `EXEMPLAR_API_USERNAME` / `EXEMPLAR_API_PASSWORD` = override the hardcoded credentials if set.

**4. On each user machine – Java**  
- **Option A – Bundle a JRE (no user install):** Place a portable Java runtime in the same folder as your .exe so the app uses it and does not require Java on the system.  
  - Download a **portable JRE** (e.g. [Eclipse Temurin](https://adoptium.net/temurin/releases/) or [Microsoft OpenJDK](https://learn.microsoft.com/en-us/java/openjdk/download) – choose **Windows x64** and **JRE** or **JDK**).  
  - Unzip it and put the folder next to your app .exe. Rename the folder to **jre** (or **runtime**). The app looks for `jre\bin\java.exe` or `runtime\bin\java.exe`.  
  - Example layout: `Student Attendance Reporting.exe`, `jre\bin\java.exe`, …  
- **Option B – System Java:** If you do not bundle a JRE, Java must be installed on the machine and on the system PATH. The app will use `java` from PATH.

If credentials are not set, the app keeps normal SQL workflow and shows API as "Not configured".

## Current API usage in app

When a student is selected, app calls:

1. `GET /api/v1/users?roles=STUDENT&firstName={firstName}&lastName={lastName}` — find the student.
2. `GET /api/v1/users/:id/cards/summary` — get card summary for that user.

From the **cards/summary** response we currently use only: total cards, completed count, pending count, and status (e.g. APPROVED). The API may return more (e.g. qualification name, card list, due dates).

**Optional debug:** set env **EXEMPLAR_DEBUG_LIST_RESPONSES=1** to write **ExemplarApiResponseFields.txt** (each API URL and every response field as `path = value`) in the app folder when a student is selected. Not needed for normal use.

**Raw JSON only:** set **EXEMPLAR_DEBUG_SAVE_JSON=1** to write the raw cards/summary JSON to **ExemplarCardsSummary_sample.json**.

The result is shown on the main form as a profiling API status line.

## Completion endpoint support (ready for wiring)

Code support is included for:

- `PUT /api/v1/users/:id/qualifications/:qualificationId`

Valid status values:

- `ACTIVE`
- `INACTIVE`
- `COMPLETED`
- `WITHDRAWN`

This can be wired into your "student complete" action in a later step.
