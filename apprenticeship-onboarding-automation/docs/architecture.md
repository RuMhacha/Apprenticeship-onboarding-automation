# Architecture

This document describes how the pieces of the system fit together. It's
aimed at someone technical who wants to understand the design without
reading every line of code.

## Components

| Component | Role |
|---|---|
| **Learner Google Form** | Collects the apprentice's own information (personal details, employer name, course choice, uploaded ID documents). |
| **Employer Google Form** | Collects the employer/line manager's confirmation details and e-signature. |
| **Google Sheets** ("Learner Management") | Central data store. Raw form responses land in dedicated "Raw" sheets; processed records live in `Learners`, `Employers`, `Programmes`, `Enrolment Details`, `Document Logs`, and a config `Settings` sheet. |
| **Google Apps Script** | Bound to the spreadsheet. Listens for form submissions and sheet edits, and drives all the business logic. |
| **Google Docs templates** | One template per document type per programme, containing `{{PLACEHOLDER}}` tokens. |
| **Google Drive** | Stores each learner's folder structure and the generated Docs/PDFs. |

## Script structure

The original script is a single Apps Script project. For this public
repository it's split into five files along logical boundaries:

- **`config.gs`** — placeholder configuration values and `getSettingsMap_()`, which reads the same key/value pairs from the spreadsheet's `Settings` sheet at runtime (so template/folder IDs can be changed by staff without touching code).
- **`utilities.gs`** — small, stateless helpers used throughout the project (header-row indexing, date formatting, regex escaping, name/company matching).
- **`template-selection.gs`** — the conditional logic that decides which Google Docs template to use, based on a learner's programme and the document being generated.
- **`form-processing.gs`** — everything triggered by a form submission or a sheet edit: creating/matching Learner and Employer records, recalculating derived fields, and reacting to status changes.
- **`document-generation.gs`** — the three document-generation workflows (Agreement, Enrolment Pack, Learner Diagnostics) and enrolment finalisation.

## Data flow

See [`workflow.md`](./workflow.md) for the step-by-step flow from form
submission to generated documents.

## Design choices worth noting

- **Settings sheet over hardcoded config.** All Drive/Docs IDs are stored
  in a `Settings` sheet and read at runtime, rather than hardcoded in the
  script. This means non-technical staff can update a template ID without
  touching Apps Script.
- **Idempotent document generation.** Each generation workflow checks
  whether a document already exists for that learner before creating a
  new one, so re-running a trigger (or a manual re-click) doesn't create
  duplicates.
- **Locking around writes.** Form-submission and document-generation
  functions acquire a script lock before reading/writing the spreadsheet,
  since multiple form submissions or edits could otherwise race.
- **Duplicate-learner protection.** Before creating a new Learner record,
  the script checks first name, surname, email, and date of birth against
  existing records to catch accidental double submissions.
