# SAP Document Archiving Automation

Automates the retrieval and archiving of supporting PDF documents for
accounting entries posted in SAP, replacing a manual monthly process
of creating folders and exporting PDFs one by one.

> **Note:** this is a sanitized portfolio copy. All company-specific
> identifiers (server names, paths, company codes, employee usernames)
> have been replaced with environment variables or generic
> placeholders — see `.env.example`. No real business data is
> included anywhere in this repository.

## Problem

A finance team manually archived supporting documents for every
accounting entry posted each month:

1. Pull a list of document numbers from SAP.
2. Manually create a folder per document on a shared network drive.
3. Manually export each document's PDF from SAP into its folder.
4. Wait for supporting files from teammates via email to complete
   each folder.

Multiple preparers each export their own list from SAP, so the raw
data can end up spread across several files, sitting in the same
folder for months at a time, mixed together.

## What this does

- Reads **all** export files in a folder, filters by content (year,
  period, cancellation flag) rather than filename, and merges them
  into a single deduplicated list — so it doesn't matter which files
  are present or how they're named.
- Compares that list against what already exists on the network
  share, checking for the folder **and** its PDF — a folder without
  its PDF is treated as incomplete and retried automatically.
- Creates missing folders and downloads the corresponding PDFs
  directly from SAP.
- Deliberately leaves attaching supporting documents (received by
  email) as a manual step — that data has no reliable source to
  automate from.
- Generates a CSV status report on every run.

## What this deliberately does *not* automate

Some tedious tasks are automated on purpose only up to the point
where automation stops being safe or reliable. Attaching third-party
supporting documents received by email was left manual because no
system-of-record exists for pairing an inbox attachment to a
document number — automating that guess would be riskier than the
time it saves.

## Safety design

This script has write access to a shared, multi-user network
location, so its design treats that as the primary constraint:

- No delete operations anywhere in the codebase.
- Every create/write is preceded by an existence check.
- `DRY_RUN=true` (the default) runs the full analysis and prints
  exactly what *would* happen, with zero side effects.
- A live run still requires typing an explicit confirmation before
  touching anything real.
- Each record is processed in its own `try/except`, so one failure
  doesn't abort the whole batch.

## Tech notes

- SAP GUI Scripting via COM automation (`pywin32`), including
  recovering from an inconsistent post-reset "welcome" screen.
- Native Windows dialog automation (`pywinauto`) with a two-backend
  fallback (`uia` then `win32`) for a save dialog that behaves
  differently across environments.
- Idempotent/resumable execution based on actual output-file
  existence, not just folder presence.

## Requirements

```
pip install pywin32 pywinauto python-dotenv
```

Requires an already-open, logged-in SAP GUI session with GUI
Scripting enabled (Options > Accessibility & Scripting > Scripting),
and a printer named "Microsoft Print to PDF" available in Windows.

## Configuration

Copy `.env.example` to `.env` and fill in your own values (network
path, SAP company code, favorite node IDs, etc.), then load them
before running — e.g. with `python-dotenv`, or by exporting them in
your shell.

## Usage

```bash
python sap_document_archiving_automation.py
```

Runs in dry-run mode by default. Set `DRY_RUN=false` in your
environment once you've reviewed the dry-run output and are ready
for a live run.
