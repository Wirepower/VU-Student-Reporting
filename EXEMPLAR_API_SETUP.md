# Exemplar API integration notes

The app now includes read-only student profiling lookup against Exemplar APIs.

## Configuration (easiest deployment)

The app reads profiling settings from `My.Settings` first (built into install), then environment variables as fallback.

Settings used:

- `ExemplarApiToken`
- `ExemplarApiBaseUrl` (default: `https://api.profiling.exemplarsystems.com.au`)
- `ExemplarQualificationId`

For easiest rollout to non-technical users, set these values in your release build so users only install and run.

## Current API usage in app

When a student is selected, app calls:

1. `GET /api/v1/users?roles=STUDENT&firstName={firstName}&lastName={lastName}`
2. `GET /api/v1/users/:id/cards/summary`

The result is shown on the main form as a profiling API status line.

On the **Student Units** form, the **Refresh Profiling %** button calls:

1. `GET /api/v1/users/:id/qualifications/:qualificationId`
2. `GET /api/v1/users/:id/qualifications/:qualificationId/units/:unitCode/progression/cards`

It appends per-unit profiling percentages/cards to each unit checkbox text without changing SQL checkbox state.

## Completion endpoint support (ready for wiring)

Code support is included for:

- `PUT /api/v1/users/:id/qualifications/:qualificationId`

Valid status values:

- `ACTIVE`
- `INACTIVE`
- `COMPLETED`
- `WITHDRAWN`

This can be wired into your "student complete" action in a later step.
