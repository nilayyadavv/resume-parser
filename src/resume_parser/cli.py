from __future__ import annotations

import argparse
import sys
from pathlib import Path

from resume_parser.extract import ParsedResume, collect_resume_files, parse_path
from resume_parser.sheets import (
    DEFAULT_SHEET_NAME,
    CredentialsNotFoundError,
    resolve_credentials,
    write_csv,
    write_google_sheet,
)


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        print(f"error: not a directory: {folder}", file=sys.stderr)
        return 2

    files = collect_resume_files(folder, recursive=args.recursive)
    if not files:
        print(f"error: no PDF or DOCX files in {folder}", file=sys.stderr)
        return 2

    results: list[ParsedResume] = []
    for path in files:
        result = parse_path(path)
        results.append(result)
        print(f"{path.name}: {result.status}  {result.name or '(no name)'}")

    csv_path = args.csv.expanduser().resolve()
    write_csv(csv_path, results)
    print(f"wrote CSV: {csv_path}")

    try:
        credentials = resolve_credentials(args.credentials)
    except CredentialsNotFoundError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if credentials is None:
        print(
            "no Google credentials found; skipped Sheets upload. "
            "Set GOOGLE_APPLICATION_CREDENTIALS or pass --credentials, "
            "then import the CSV via File → Import in Google Sheets."
        )
        return 0

    try:
        url = write_google_sheet(
            results,
            credentials_path=credentials,
            spreadsheet_id=args.spreadsheet_id,
            sheet_name=args.sheet_name,
        )
    except Exception as exc:  # noqa: BLE001 - CLI should not traceback on API errors
        print(f"error: Google Sheets write failed: {exc}", file=sys.stderr)
        print(f"CSV is still available at {csv_path}", file=sys.stderr)
        return 1

    print(f"wrote Google Sheet: {url}")
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="resume-parser",
        description=(
            "Parse PDF and DOCX resumes and write name, email, GitHub, "
            "and LinkedIn into a Google Sheet (CSV fallback without credentials)."
        ),
    )
    parser.add_argument(
        "folder",
        type=Path,
        help="Directory containing resume PDF/DOCX files",
    )
    parser.add_argument(
        "--spreadsheet-id",
        help="Existing Google Spreadsheet ID. Creates a new spreadsheet if omitted.",
    )
    parser.add_argument(
        "--sheet-name",
        default=DEFAULT_SHEET_NAME,
        help=f"Worksheet title (default: {DEFAULT_SHEET_NAME})",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=Path("parsed_resumes.csv"),
        help="Local CSV path (default: parsed_resumes.csv)",
    )
    parser.add_argument(
        "--credentials",
        help="Path to a Google service account JSON file "
        "(defaults to GOOGLE_APPLICATION_CREDENTIALS)",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Scan subfolders for resumes",
    )
    return parser


if __name__ == "__main__":
    raise SystemExit(main())
