# Exemplar API integration notes

The app now includes read-only student profiling lookup against Exemplar APIs.

## Environment variables

Set these on client machines (or user session) before launching the app:

- `EXEMPLAR_API_TOKEN` = Bearer token value (without `Bearer ` prefix)
- `EXEMPLAR_API_BASE_URL` = optional override  
  default: `https://api.profiling.exemplarsystems.com.au`

If token is missing, the app keeps normal SQL workflow and shows API as "Not configured".

## Current API usage in app

When a student is selected, app calls:

1. `GET /api/v1/users?roles=STUDENT&firstName={firstName}&lastName={lastName}`
2. `GET /api/v1/users/:id/cards/summary`

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
