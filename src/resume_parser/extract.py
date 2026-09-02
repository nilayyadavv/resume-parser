from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from pathlib import Path

from docx import Document
from pypdf import PdfReader

MAX_FILE_BYTES = 10 * 1024 * 1024
SUPPORTED_SUFFIXES = {".pdf", ".docx"}

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
GITHUB_RE = re.compile(
    r"(?:https?://)?(?:www\.)?github\.com/([A-Za-z0-9](?:[A-Za-z0-9-]{0,37}[A-Za-z0-9])?)(?:/[^\s]*)?",
    re.I,
)
LINKEDIN_RE = re.compile(
    r"(?:https?://)?(?:www\.)?linkedin\.com/(in|pub)/([A-Za-z0-9\-_%]+)(?:/[^\s]*)?",
    re.I,
)
PHONE_RE = re.compile(r"\d{3}[-.\s]\d{3}[-.\s]\d{4}")
SKIP_NAME_RE = re.compile(
    r"resume|curriculum|vitae|\bcv\b|contact|email|phone|profile|objective|"
    r"summary|experience|education|skills|projects|github|linkedin|http",
    re.I,
)
GITHUB_RESERVED = {
    "about",
    "blog",
    "collections",
    "customer-stories",
    "enterprise",
    "events",
    "explore",
    "features",
    "github-copilot",
    "issues",
    "login",
    "marketplace",
    "new",
    "notifications",
    "orgs",
    "organizations",
    "pricing",
    "pulls",
    "search",
    "security",
    "settings",
    "signup",
    "solutions",
    "sponsors",
    "topics",
    "trending",
}

STATUS_OK = "ok"
STATUS_NO_LINKS = "no GitHub or LinkedIn found"
STATUS_UNSUPPORTED = "unsupported file type"
STATUS_TOO_LARGE = "file too large (over 10MB)"
STATUS_EMPTY = "no extractable text"
STATUS_FAILED = "could not read file"


@dataclass
class ParsedResume:
    name: str = ""
    email: str = ""
    github: str = ""
    linkedin: str = ""
    source_file: str = ""
    status: str = STATUS_OK

    def as_row(self) -> list[str]:
        return [
            self.name,
            self.email,
            self.github,
            self.linkedin,
            self.source_file,
            self.status,
        ]

    def as_dict(self) -> dict[str, str]:
        return asdict(self)


def parse_path(path: Path) -> ParsedResume:
    result = ParsedResume(source_file=path.name)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        result.status = STATUS_UNSUPPORTED
        return result
    try:
        size = path.stat().st_size
    except OSError as exc:
        result.status = f"{STATUS_FAILED}: {exc}"
        return result
    if size > MAX_FILE_BYTES:
        result.status = STATUS_TOO_LARGE
        return result
    try:
        text = extract_text(path)
    except Exception as exc:  # noqa: BLE001 - surface any reader failure in the row
        result.status = f"{STATUS_FAILED}: {exc}"
        return result
    return parse_text(text, filename=path.name)


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return _pdf_text(path)
    if suffix == ".docx":
        return _docx_text(path)
    raise ValueError(f"Unsupported file type: {suffix}")


def parse_text(text: str, filename: str = "") -> ParsedResume:
    normalized = _normalize_text(text)
    email = _first_email(normalized)
    github = _first_github(normalized)
    linkedin = _first_linkedin(normalized)
    name = _first_name(normalized, filename)
    if not normalized.strip():
        status = STATUS_EMPTY
    elif not github and not linkedin:
        status = STATUS_NO_LINKS
    else:
        status = STATUS_OK
    return ParsedResume(
        name=name,
        email=email,
        github=github,
        linkedin=linkedin,
        source_file=filename,
        status=status,
    )


def collect_resume_files(folder: Path, recursive: bool = False) -> list[Path]:
    if recursive:
        files = [p for p in folder.rglob("*") if p.is_file()]
    else:
        files = [p for p in folder.iterdir() if p.is_file()]
    return sorted(
        p
        for p in files
        if p.suffix.lower() in SUPPORTED_SUFFIXES and not p.name.startswith(".")
    )


def _pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages)


def _docx_text(path: Path) -> str:
    document = Document(str(path))
    parts: list[str] = []
    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            parts.append(paragraph.text)
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                parts.append("\n".join(cells))
    return "\n".join(parts)


def _normalize_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def _first_email(text: str) -> str:
    match = EMAIL_RE.search(text)
    return match.group(0) if match else ""


def _first_github(text: str) -> str:
    for match in GITHUB_RE.finditer(text):
        username = match.group(1)
        if username.lower() in GITHUB_RESERVED:
            continue
        return f"https://github.com/{username}"
    return ""


def _first_linkedin(text: str) -> str:
    match = LINKEDIN_RE.search(text)
    if not match:
        return ""
    kind, slug = match.group(1).lower(), match.group(2)
    return f"https://www.linkedin.com/{kind}/{slug}"


def _first_name(text: str, filename: str) -> str:
    for raw_line in text.splitlines()[:16]:
        line = raw_line.strip(" |-•\t")
        if not line or EMAIL_RE.search(line) or PHONE_RE.search(line):
            continue
        if SKIP_NAME_RE.search(line):
            continue
        words = line.split()
        if 2 <= len(words) <= 5 and len(line) <= 60:
            return line.title() if line.isupper() else line
        if len(words) == 1 and line[:1].isupper() and line.isalpha() and len(line) >= 2:
            return line
    return _name_from_filename(filename)


def _name_from_filename(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"(?i)[_-]?resume[_-]?", " ", stem)
    stem = stem.replace("_", " ").replace("-", " ")
    stem = re.sub(r"\s+", " ", stem).strip()
    return stem.title() if stem else ""
