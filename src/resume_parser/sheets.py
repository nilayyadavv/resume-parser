from __future__ import annotations

import csv
import os
from pathlib import Path

from resume_parser.extract import ParsedResume

HEADERS = ["Name", "Email", "GitHub", "LinkedIn", "Source file", "Status"]

DEFAULT_SPREADSHEET_TITLE = "Resume Parser Output"
DEFAULT_SHEET_NAME = "Resumes"


def write_csv(path: Path, results: list[ParsedResume]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "Name": result.name,
                    "Email": result.email,
                    "GitHub": result.github,
                    "LinkedIn": result.linkedin,
                    "Source file": result.source_file,
                    "Status": result.status,
                }
            )
    return path


class CredentialsNotFoundError(FileNotFoundError):
    """Raised when an explicit credentials path does not exist."""


def resolve_credentials(credentials: str | None) -> Path | None:
    explicit = credentials is not None
    raw = credentials or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not raw:
        return None
    path = Path(raw).expanduser()
    if path.is_file():
        return path
    if explicit:
        raise CredentialsNotFoundError(f"credentials file not found: {path}")
    return None


def write_google_sheet(
    results: list[ParsedResume],
    *,
    credentials_path: Path,
    spreadsheet_id: str | None = None,
    sheet_name: str = DEFAULT_SHEET_NAME,
) -> str:
    import gspread

    client = gspread.service_account(filename=str(credentials_path))
    if spreadsheet_id:
        spreadsheet = client.open_by_key(spreadsheet_id)
    else:
        spreadsheet = client.create(DEFAULT_SPREADSHEET_TITLE)

    worksheet = _get_or_create_worksheet(spreadsheet, sheet_name)
    rows = [HEADERS] + [result.as_row() for result in results]
    worksheet.clear()
    worksheet.update(rows, value_input_option="USER_ENTERED")
    worksheet.format("A1:F1", {"textFormat": {"bold": True}})
    worksheet.freeze(rows=1)
    try:
        worksheet.columns_auto_resize(0, 5)
    except Exception:
        pass
    return spreadsheet.url


def _get_or_create_worksheet(spreadsheet, sheet_name: str):
    import gspread

    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        row_count = 100
        return spreadsheet.add_worksheet(title=sheet_name, rows=row_count, cols=len(HEADERS))
