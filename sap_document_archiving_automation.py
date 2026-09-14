"""
SAP DOCUMENT ARCHIVING AUTOMATION
===================================

PORTFOLIO / DEMO VERSION -- all company-specific identifiers (server
names, paths, company codes, employee usernames) have been replaced
with environment variables or generic placeholders. This is a
sanitized copy of a production script; no real business data is
included.

PROBLEM
-------
A finance team manually archives supporting PDF documents for every
accounting document (journal entry) posted each month. The process
was: pull a list of document numbers from SAP, manually create a
folder per document on a shared network drive, manually export each
document's PDF from SAP into its folder, then wait for supporting
files from other team members via email to complete each folder.

Multiple preparers each pull their own list from SAP, so the source
data arrives as several separate export files that may sit in the
same folder for months at a time, mixed together.

This script automates the tedious, error-prone parts (folder creation
and PDF retrieval) while deliberately leaving the human-dependent part
(attaching supporting documents received by email) as a manual step.

WHAT THIS DEMONSTRATES
-----------------------
  - SAP GUI Scripting via COM automation (win32com), including
    navigating a multi-window print flow and recovering from an
    unpredictable "welcome" screen that SAP sometimes shows after a
    session reset.
  - Windows native dialog automation (pywinauto) with a two-backend
    fallback strategy (UI Automation, then legacy Win32) for dialogs
    that behave inconsistently across environments.
  - A "content over filename" parsing strategy: rather than relying on
    file naming conventions, records are filtered using fields inside
    the export itself (year, period, cancellation flag), so the
    exact set of source files present in a folder never matters.
  - Idempotent, resumable execution: completion is determined by
    checking for the actual expected output file, not just a folder's
    existence -- so a partially failed run is retried automatically
    on the next execution rather than silently left incomplete.
  - Defense-in-depth safety design for a script with write access to
    a shared, multi-user network location:
      * No delete operations anywhere in the codebase.
      * Existence checks before every create/write operation.
      * A dry-run mode that performs the full read-only analysis and
        prints exactly what would happen, with zero side effects.
      * An explicit typed confirmation gate before any real write.
      * Per-record error isolation so one failure doesn't abort a
        multi-record batch.
  - A CSV status report generated on every run (dry or real) for
    downstream visibility into what's outstanding and who it belongs
    to.

CONFIGURATION
-------------
All environment-specific values are read from environment variables.
Copy `.env.example` to `.env` and fill in your own values -- it is
loaded automatically on startup via `python-dotenv`.

REQUIREMENTS
------------
  pip install -r requirements.txt

Requires an already-open, logged-in SAP GUI session with GUI Scripting
enabled (Options > Accessibility & Scripting > Scripting), and a
printer named "Microsoft Print to PDF" available in Windows.
"""

import csv
import glob
import os
import time

import win32com.client as win32
from dotenv import load_dotenv
from pywinauto import Application

load_dotenv()


# =============================================================================
# CONFIGURATION -- populate via environment variables, never hardcode
# real company data here.
# =============================================================================

DRY_RUN = os.environ.get("DRY_RUN", "true").lower() != "false"
# True  -> simulation only, zero filesystem/SAP writes.
# False -> live run, will create folders and download PDFs.

COMPANY_CODE = os.environ.get("SAP_COMPANY_CODE", "XXXX")
YEAR = os.environ.get("TARGET_YEAR", "2026")
MONTH = os.environ.get("TARGET_MONTH", "01")  # zero-padded, e.g. "01".."12"

# Optional: restrict processing to a single preparer's records.
# Leave as None (unset) to process everyone found in the source files.
USER_FILTER = os.environ.get("USER_FILTER") or None

DOCUMENT_ARCHIVE_ROOT = os.environ.get(
    "DOCUMENT_ARCHIVE_ROOT",
    r"\\your-file-server\shared\accounting\document_archive\{year}\{company}\SA",
).format(year=YEAR, company=COMPANY_CODE)

SOURCE_EXPORT_FOLDER = os.environ.get("SOURCE_EXPORT_FOLDER", r"C:\temp\sap_document_lists")

SAP_DOCUMENT_FAVORITE_NODE = os.environ.get("SAP_DOCUMENT_FAVORITE_NODE", "F00002")

# Column layout of the tab-delimited SAP export (0-indexed). Adjust these
# if your own export's field order/count differs -- this is the one part
# of the script that is tied to a specific SAP list layout rather than
# being generic.
COL_COMPANY = 1
COL_DOC_NUMBER = 3
COL_DOC_YEAR = 4
COL_DOC_PERIOD = 5
COL_PREPARER = 11
COL_CANCELLATION_REF = 20
MIN_EXPECTED_COLUMNS = 12


# =============================================================================
# STEP 1: Read the document-number list exported from SAP
# =============================================================================

def get_source_records() -> dict:
    """
    Reads EVERY file present in SOURCE_EXPORT_FOLDER (this folder may
    contain exports from several months/preparers mixed together) and
    returns a dict {document_number: preparer} -- limited to rows that
    match exactly the configured YEAR and MONTH.

    The source format is tab-delimited, structured as:
    (blank) | Company | Reference | Doc.Number | Year | Period | ...

    A row is valid when: Company == COMPANY_CODE, Year == YEAR,
    Period == MONTH (as an integer), Doc.Number is purely numeric, and
    the cancellation-reference column is empty (a non-empty value
    means the document was voided and doesn't need archiving).
    """
    files = [
        f for f in glob.glob(os.path.join(SOURCE_EXPORT_FOLDER, "*"))
        if os.path.isfile(f) and not os.path.basename(f).startswith("report_")
    ]
    if not files:
        raise RuntimeError(f"No files found in {SOURCE_EXPORT_FOLDER}")

    target_month = int(MONTH)
    target_year = str(int(YEAR))

    records = {}

    for file_path in files:
        try:
            # latin-1 because classic SAP exports aren't UTF-8 and
            # contain special characters that would fail otherwise.
            with open(file_path, encoding="latin-1") as f:
                for line in f:
                    parts = line.split("\t")
                    if len(parts) < MIN_EXPECTED_COLUMNS:
                        continue

                    company = parts[COL_COMPANY].strip()
                    doc_number = parts[COL_DOC_NUMBER].strip()
                    doc_year = parts[COL_DOC_YEAR].strip()
                    doc_period = parts[COL_DOC_PERIOD].strip()
                    preparer = parts[COL_PREPARER].strip()
                    cancellation_ref = (
                        parts[COL_CANCELLATION_REF].strip()
                        if len(parts) > COL_CANCELLATION_REF else ""
                    )

                    if company != COMPANY_CODE or not doc_number.isdigit():
                        continue
                    if doc_year != target_year:
                        continue
                    try:
                        if int(doc_period) != target_month:
                            continue
                    except ValueError:
                        continue
                    if cancellation_ref != "":
                        continue

                    records[doc_number] = preparer

        except (UnicodeDecodeError, PermissionError):
            print(f"  (skipped '{file_path}', could not read as text)")
            continue

    print(f"Read {len(files)} file(s) in {SOURCE_EXPORT_FOLDER}")
    print(f"Document numbers found for {MONTH}/{YEAR}: {len(records)}")

    return records


# =============================================================================
# STEP 2: Compare against what already exists (read-only)
# =============================================================================

def get_completed_documents() -> set:
    """
    A document is considered 'complete' only if BOTH its folder AND its
    PDF exist. A folder without its PDF (e.g. left behind by a
    partially failed run) does NOT count as complete -- so it gets
    retried automatically on the next run instead of being silently
    orphaned.
    """
    if not os.path.isdir(DOCUMENT_ARCHIVE_ROOT):
        raise RuntimeError(f"Cannot access: {DOCUMENT_ARCHIVE_ROOT}")

    completed = set()
    for name in os.listdir(DOCUMENT_ARCHIVE_ROOT):
        folder_path = os.path.join(DOCUMENT_ARCHIVE_ROOT, name)
        if os.path.isdir(folder_path) and os.path.isfile(expected_pdf_path(name)):
            completed.add(name)
    return completed


def expected_pdf_path(doc_number: str) -> str:
    """Builds the expected PDF path for a given document number (does not check existence)."""
    filename = f"DOC_{COMPANY_CODE}_{YEAR}_{doc_number}.pdf"
    return os.path.join(DOCUMENT_ARCHIVE_ROOT, doc_number, filename)


def generate_report(source_records: dict, completed: set) -> str:
    """Writes a CSV status report: document number, preparer, status, expected PDF path."""
    report_filename = f"report_{COMPANY_CODE}_{YEAR}_{MONTH}.csv"
    report_path = os.path.join(SOURCE_EXPORT_FOLDER, report_filename)

    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Document Number", "Preparer", "Status", "Expected PDF Path"])
        for doc_number in sorted(source_records):
            preparer = source_records[doc_number]
            status = "Complete" if doc_number in completed else "Pending"
            writer.writerow([doc_number, preparer, status, expected_pdf_path(doc_number)])

    print(f"\nReport generated: {report_path}")
    return report_path


# =============================================================================
# STEP 3: SAP -- open the document and trigger PDF printing
# =============================================================================

def connect_to_sap():
    """Attaches to the first already-open SAP GUI session on this machine."""
    sap_gui_auto = win32.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def open_document_and_print(session, doc_number: str) -> None:
    """
    Opens the given document and drives the print flow through to the
    point where the native "Save Print Output As" dialog appears.

    The print("[diag] ...") lines are checkpoints: if the script fails,
    the LAST checkpoint printed before the error tells you exactly
    which step it broke on -- valuable when debugging screen-scraping
    style automation that has no stack trace pointing at "the wrong
    SAP screen".
    """
    print("    [diag] Returning to the initial screen...")
    # "/n" in the transaction-code field, followed by Enter, is the
    # equivalent of "close whatever you're doing and go back to the
    # main menu" -- necessary because after the first document, SAP is
    # left on the previous document's display screen.
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    # Sometimes, after "/n", SAP shows a blank welcome screen with a
    # "Start SAP Easy Access" button instead of going directly to the
    # favorites tree. If it appears, it needs to be clicked first.
    try:
        session.findById("wnd[0]/usr/btnSTARTBUTTON").press()
        print("    [diag] Welcome screen appeared, clicked through it.")
        time.sleep(1)
    except Exception:
        pass  # that screen didn't appear this time, continue normally

    print("    [diag] Opening the document display favorite...")
    session.findById(
        "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"
    ).doubleClickNode(SAP_DOCUMENT_FAVORITE_NODE)
    time.sleep(1)

    print("    [diag] Entering document number, company and year...")
    session.findById("wnd[0]/usr/txtRF05L-BELNR").text = doc_number
    session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = COMPANY_CODE
    session.findById("wnd[0]/usr/txtRF05L-GJAHR").text = YEAR
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    print("    [diag] Navigating the print menu...")
    session.findById("wnd[0]/mbar/menu[0]/menu[5]").select()
    time.sleep(0.5)
    session.findById("wnd[0]/mbar/menu[0]/menu[0]").select()
    time.sleep(1)

    print("    [diag] Configuring local printer...")
    session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").text = "locl"
    session.findById(
        "wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM"
    ).key = "X"
    time.sleep(0.5)

    print("    [diag] Confirming intermediate dialog...")
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    time.sleep(0.5)

    print("    [diag] Triggering print...")
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    time.sleep(2)
    print("    [diag] Native save dialog should now be visible.")


# =============================================================================
# STEP 4: Write the path into the native "Save As" dialog
# =============================================================================

# pywinauto's type_keys() treats these characters as modifier/special-key
# markers (e.g. "%" = Alt, "^" = Ctrl, "~" = Enter) rather than literal
# text. Wrapping each one in its own {...} escapes it back to a literal
# keystroke -- without this, a company code or archive path that happens
# to contain one of them would silently type the wrong thing into the
# save dialog instead of raising an error.
_TYPE_KEYS_SPECIAL_CHARS = "+^%~(){}"


def _escape_for_type_keys(text: str) -> str:
    return "".join(f"{{{ch}}}" if ch in _TYPE_KEYS_SPECIAL_CHARS else ch for ch in text)


def save_pdf(destination_path: str) -> None:
    """
    Writes the full destination path into Windows' native "Save Print
    Output As" dialog and confirms with Enter.

    Tries two pywinauto backends: "uia" (modern, can target the
    filename field by its visible label) and "win32" (legacy, falls
    back to the Alt+N keyboard shortcut that jumps to the filename
    field in any standard Windows save dialog). Different environments
    have shown one or the other to be more reliable.
    """
    dlg = None
    backend_used = None

    for backend in ("uia", "win32"):
        try:
            app = Application(backend=backend).connect(
                title_re=".*Save.*Print.*", timeout=5
            )
            dlg = app.window(title_re=".*Save.*Print.*")
            backend_used = backend
            break
        except Exception:
            continue

    if dlg is None:
        raise RuntimeError("The native 'Save Print Output As' dialog did not appear.")

    dlg.set_focus()
    time.sleep(0.3)

    if backend_used == "uia":
        filename_field = dlg.child_window(title="File name:", control_type="ComboBox")
        filename_field.set_focus()
    else:
        dlg.type_keys("%n", pause=0.2)  # Alt+N jumps to the filename field
    time.sleep(0.3)

    dlg.type_keys("^a", pause=0.1)
    dlg.type_keys("{DEL}", pause=0.1)
    dlg.type_keys(_escape_for_type_keys(destination_path), with_spaces=True, pause=0.01)
    time.sleep(0.3)

    dlg.type_keys("~", pause=0.1)  # Enter -> default "Save" button
    time.sleep(1.5)


# =============================================================================
# MAIN ORCHESTRATION
# =============================================================================

def main():
    print(f"Mode: {'DRY RUN (no changes)' if DRY_RUN else 'LIVE RUN'}")

    source_records = get_source_records()

    if USER_FILTER is not None:
        before = len(source_records)
        source_records = {
            doc_number: preparer for doc_number, preparer in source_records.items()
            if preparer == USER_FILTER
        }
        print(f"User filter active: '{USER_FILTER}' -- {len(source_records)}/{before} records match.")

    completed = get_completed_documents()
    pending = [n for n in source_records if n not in completed]

    print(f"\nTotal in source list: {len(source_records)}")
    print(f"Already complete (folder + PDF): {len(completed)}")
    print(f"Pending/incomplete: {len(pending)}")

    if DRY_RUN:
        print("\n--- Would create these folders and download these PDFs ---")
        for doc_number in pending:
            print(f"  {doc_number} ({source_records[doc_number]}) -> {expected_pdf_path(doc_number)}")
        print("\n(Nothing was touched -- this is a simulation report only)")
        generate_report(source_records, completed)
        return

    # --- Live run only from here on ---

    print(f"\nAbout to create {len(pending)} folder(s) and download their PDFs to:")
    print(f"  {DOCUMENT_ARCHIVE_ROOT}")
    answer = input("\nType YES (all caps) to proceed, anything else to cancel: ")
    if answer != "YES":
        print("Cancelled. Nothing was touched.")
        return

    session = connect_to_sap()
    created, skipped, errors = 0, 0, 0

    for doc_number in pending:
        preparer = source_records[doc_number]
        folder_path = os.path.join(DOCUMENT_ARCHIVE_ROOT, doc_number)
        pdf_path = expected_pdf_path(doc_number)

        print(f"\n--- Processing {doc_number} (preparer: {preparer}) ---")

        # Safety lock: never overwrite an existing PDF, even if it was
        # already filtered out above -- this is intentional redundancy.
        if os.path.exists(pdf_path):
            print(f"  {doc_number}: PDF already exists, skipping.")
            skipped += 1
            continue

        # Only create the folder if it doesn't already exist.
        if not os.path.isdir(folder_path):
            os.makedirs(folder_path)
            print(f"  {doc_number}: folder created.")

        try:
            open_document_and_print(session, doc_number)
            save_pdf(pdf_path)

            if os.path.exists(pdf_path):
                print(f"  {doc_number}: PDF downloaded successfully.")
                created += 1
            else:
                print(f"  {doc_number}: PDF not found after saving, needs manual review.")
                errors += 1
        except Exception as e:
            # Isolating the failure here lets the loop continue with
            # the next document instead of aborting the whole batch.
            print(f"  {doc_number}: error -- {e}")
            errors += 1

    print(f"\nSummary: {created} new PDFs, {skipped} skipped (already existed), {errors} errors.")

    final_completed = get_completed_documents()
    generate_report(source_records, final_completed)


if __name__ == "__main__":
    main()
