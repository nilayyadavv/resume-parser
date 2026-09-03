# Resume Parser

Parse a folder of PDF and DOCX resumes and write **name, email, GitHub, and LinkedIn** into a Google Sheet. If Google credentials are missing, the same rows are written to a CSV you can import into Sheets.

## What it extracts

| Column | How it is found |
| --- | --- |
| Name | First prominent name line, including first/last split across lines |
| Email | First real email (placeholder `yourname@example.com` values ignored) |
| GitHub | First `github.com/<user>` URL in text or clickable links |
| LinkedIn | First `linkedin.com/in/...` or `linkedin.com/pub/...` URL in text or clickable links |
| Source file | Original filename |
| Status | `ok`, `no GitHub or LinkedIn found`, or a read error |

It does not call the GitHub or LinkedIn APIs. Links come from visible text plus PDF/DOCX hyperlinks.

## Install

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```bash
uv sync
```

## Parse a folder of resumes

```bash
uv run resume-parser ./resumes
```

That writes `parsed_resumes.csv` in the current directory. Sample resumes are in [`resumes/`](resumes/).

Useful flags:

```bash
# Recurse into subfolders
uv run resume-parser ./resumes --recursive

# Choose the CSV path
uv run resume-parser ./resumes --csv ./out/candidates.csv

# Write into an existing Google Sheet (share it with the service account)
uv run resume-parser ./resumes --spreadsheet-id YOUR_SHEET_ID --sheet-name Resumes

# Explicit credentials file
uv run resume-parser ./resumes --credentials ./credentials.json
```

## Google Sheets setup

1. Create a Google Cloud project and enable the **Google Sheets API** and **Google Drive API**.
2. Create a **service account** and download its JSON key.
3. Point the parser at that file:

   ```bash
   export GOOGLE_APPLICATION_CREDENTIALS=/path/to/credentials.json
   ```

4. Either:
   - omit `--spreadsheet-id` to create a new spreadsheet titled `Resume Parser Output` (the service account becomes the owner; share that sheet with your own Google account from Drive), or
   - create a sheet yourself, share it with the service account email as Editor, and pass `--spreadsheet-id`.

Without credentials the parser still succeeds: import `parsed_resumes.csv` in Google Sheets with **File → Import**.

Do not commit `credentials.json`. It is gitignored.

## Tests and sample files

```bash
uv run pytest
uv run python scripts/generate_samples.py
```

## GitHub repository visibility

Start this project as a **private** GitHub repository. When you want it public, use GitHub **Settings → General → Danger zone → Change repository visibility**.
