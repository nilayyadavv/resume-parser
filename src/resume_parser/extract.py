from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from pypdf import PdfReader

MAX_FILE_BYTES = 10 * 1024 * 1024
SUPPORTED_SUFFIXES = {".pdf", ".docx"}

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
GITHUB_RE = re.compile(
    r"(?:https?://)?(?:www\.)?github\.com/([A-Za-z0-9](?:[A-Za-z0-9-]{0,37}[A-Za-z0-9])?)(?:/[^\s]*)?",
    re.I,
)
LINKEDIN_RE = re.compile(
    r"(?:https?://)?(?:(?:www|[a-z]{2})\.)?linkedin\.com/(in|pub)/([A-Za-z0-9\-_%]+)(?:/[^\s]*)?",
    re.I,
)
PHONE_RE = re.compile(
    r"(?:\+?\d{1,3}[\s.-]*)?(?:\(?\d{3}\)?[\s.-]*)\d{3}[\s.-]*\d{4}"
)
SKIP_NAME_RE = re.compile(
    r"resume|curriculum|vitae|\bcv\b|contact|email|phone|profile|objective|"
    r"summary|experience|education|skills|projects|github|linkedin|http|"
    r"website|portfolio|address|location|references|certificat|languages|"
    r"awards|interests|hobbies|volunteer|publication|coursework|technical|"
    r"professional|bachelor|master|university|college|company|intern",
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
    "gist",
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
PLACEHOLDER_USERS = {
    "account",
    "example",
    "me",
    "name",
    "placeholder",
    "profile",
    "user",
    "username",
    "your-name",
    "your-username",
    "yourname",
    "yourusername",
}
PLACEHOLDER_EMAIL_LOCALS = {
    "email",
    "firstname.lastname",
    "me",
    "name",
    "placeholder",
    "test",
    "user",
    "username",
    "youremail",
    "yourname",
}
PLACEHOLDER_EMAIL_DOMAINS = {
    "domain.com",
    "email.com",
    "example.com",
    "example.net",
    "example.org",
    "test.com",
    "yourdomain.com",
}
PLACEHOLDER_NAMES = {
    "candidate name",
    "first last",
    "first name",
    "full name",
    "jane doe",
    "john doe",
    "last name",
    "student name",
    "your name",
}
NAME_STOPWORDS = {
    "activity",
    "address",
    "analyst",
    "available",
    "award",
    "awards",
    "bachelor",
    "city",
    "college",
    "company",
    "computer",
    "contact",
    "course",
    "coursework",
    "curriculum",
    "cv",
    "data",
    "degree",
    "describe",
    "developer",
    "director",
    "document",
    "education",
    "email",
    "engineer",
    "engineering",
    "experience",
    "frameworks",
    "github",
    "gpa",
    "graduation",
    "hobbies",
    "impact",
    "interests",
    "intern",
    "internship",
    "languages",
    "leadership",
    "linkedin",
    "location",
    "manager",
    "master",
    "month",
    "name",
    "objective",
    "phone",
    "portfolio",
    "professional",
    "profile",
    "programming",
    "project",
    "projects",
    "publication",
    "references",
    "relevant",
    "remember",
    "resume",
    "role",
    "science",
    "skills",
    "software",
    "state",
    "student",
    "summary",
    "technical",
    "template",
    "university",
    "untitled",
    "vitae",
    "volunteer",
    "website",
    "work",
    "year",
}
MONTHS = {
    "april",
    "august",
    "december",
    "february",
    "january",
    "july",
    "june",
    "march",
    "may",
    "november",
    "october",
    "september",
    "apr",
    "aug",
    "dec",
    "feb",
    "jan",
    "jul",
    "jun",
    "mar",
    "nov",
    "oct",
    "sep",
    "sept",
}
GENERIC_FILENAME_WORDS = {
    "ai",
    "aiml",
    "anaylst",
    "analyst",
    "basic",
    "copy",
    "current",
    "curriculum",
    "cv",
    "data",
    "docx",
    "draft",
    "engineer",
    "engineering",
    "final",
    "intern",
    "latest",
    "linkedin",
    "ml",
    "new",
    "pdf",
    "professional",
    "real",
    "realswe",
    "resume",
    "software",
    "swe",
    "template",
    "updated",
    "vitae",
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
        return _pdf_content(path)
    if suffix == ".docx":
        return _docx_content(path)
    raise ValueError(f"Unsupported file type: {suffix}")


def parse_text(text: str, filename: str = "") -> ParsedResume:
    normalized = _normalize_text(text)
    searchable = _url_search_text(normalized)
    email = _first_email(normalized)
    github = _first_github(searchable) or _first_github(normalized)
    linkedin = _first_linkedin(searchable) or _first_linkedin(normalized)
    name = _first_name(normalized, filename)
    if _is_blank(normalized):
        status = STATUS_EMPTY
        name = name if name else ""
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


def _pdf_content(path: Path) -> str:
    reader = PdfReader(str(path))
    pages: list[str] = []
    uris: list[str] = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
        uris.extend(_pdf_page_uris(page))
    return _join_text_and_uris("\n".join(pages), uris)


def _pdf_page_uris(page) -> list[str]:
    uris: list[str] = []
    annots = page.get("/Annots") or []
    for annot in annots:
        try:
            obj = annot.get_object() if hasattr(annot, "get_object") else annot
            action = obj.get("/A") if obj is not None else None
            raw = None
            if action is not None:
                raw = action.get("/URI")
            if raw is None and obj is not None:
                raw = obj.get("/URI")
            if raw:
                uris.append(str(raw))
        except Exception:  # noqa: BLE001 - ignore malformed annotations
            continue
    return uris


def _docx_content(path: Path) -> str:
    document = Document(str(path))
    parts: list[str] = []
    parts.extend(_docx_element_lines(document.element))
    seen = {line for line in parts}
    for section in document.sections:
        for hf in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            try:
                extra = _docx_element_lines(hf._element)
            except Exception:  # noqa: BLE001 - some section parts are missing
                continue
            for line in extra:
                if line not in seen:
                    parts.append(line)
                    seen.add(line)
    uris = _docx_uris(document)
    return _join_text_and_uris("\n".join(parts), uris)


def _docx_element_lines(element) -> list[str]:
    lines: list[str] = []
    for paragraph in element.iter(qn("w:p")):
        texts = [node.text or "" for node in paragraph.iter(qn("w:t"))]
        line = "".join(texts).strip()
        if line:
            lines.append(line)
    return lines


def _docx_uris(document: Document) -> list[str]:
    uris: list[str] = []
    parts = [document.part]
    for section in document.sections:
        for hf in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            try:
                parts.append(hf.part)
            except Exception:  # noqa: BLE001
                continue
    seen: set[str] = set()
    for part in parts:
        try:
            rels = part.rels.values()
        except Exception:  # noqa: BLE001
            continue
        for rel in rels:
            if "hyperlink" not in getattr(rel, "reltype", ""):
                continue
            target = getattr(rel, "target_ref", "") or ""
            if target and target not in seen:
                seen.add(target)
                uris.append(target)
    return uris


def _join_text_and_uris(text: str, uris: list[str]) -> str:
    extra = []
    lower = text.lower()
    for uri in uris:
        cleaned = uri.strip()
        if cleaned.lower().startswith("mailto:"):
            cleaned = cleaned[7:]
        if cleaned and cleaned.lower() not in lower:
            extra.append(cleaned)
            lower += "\n" + cleaned.lower()
    if not extra:
        return text
    return text + "\n" + "\n".join(extra)


def _normalize_text(text: str) -> str:
    text = text.replace("\x00", " ").replace("\u00a0", " ").replace("\ufeff", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def _url_search_text(text: str) -> str:
    lines = [line.strip() for line in text.splitlines()]
    merged: list[str] = []
    for line in lines:
        if merged and _incomplete_url(merged[-1]) and line:
            merged[-1] = merged[-1].rstrip("/") + "/" + line.lstrip("/")
        else:
            merged.append(line)
    joined = "\n".join(merged)
    joined = re.sub(r"(?i)/linkedinlinkedin\.com", " linkedin.com", joined)
    joined = re.sub(r"(?i)/githubgithub\.com", " github.com", joined)
    joined = re.sub(
        r"(?i)((?:https?://)?(?:www\.)?github\s*\.\s*com\s*/\s*[A-Za-z0-9-]+)",
        lambda match: re.sub(r"\s+", "", match.group(1)),
        joined,
    )
    joined = re.sub(
        r"(?i)((?:https?://)?(?:(?:www|[a-z]{2})\.)?linkedin\s*\.\s*com\s*/\s*(?:in|pub)\s*/\s*[A-Za-z0-9\-_%]+)",
        lambda match: re.sub(r"\s+", "", match.group(1)),
        joined,
    )
    return joined


def _incomplete_url(line: str) -> bool:
    return bool(
        re.search(
            r"(?i)(github\.com|linkedin\.com/(?:in|pub))/?\s*$",
            line.strip(),
        )
    )


def _is_blank(text: str) -> bool:
    stripped = re.sub(r"[\s•,.\-|_|/:;]+", "", text)
    return not any(ch.isalpha() for ch in stripped)


def _first_email(text: str) -> str:
    for match in EMAIL_RE.finditer(text):
        email = match.group(0)
        if _valid_email(email):
            return email
    return ""


def _valid_email(email: str) -> bool:
    local, _, domain = email.partition("@")
    if not local or not domain:
        return False
    if local.lower() in PLACEHOLDER_EMAIL_LOCALS:
        return False
    if local.lower() in {"you.name", "first.last"}:
        return False
    if domain.lower() in PLACEHOLDER_EMAIL_DOMAINS and local.lower().startswith("your"):
        return False
    return True


def _first_github(text: str) -> str:
    candidates: list[str] = []
    for match in GITHUB_RE.finditer(text):
        username = match.group(1)
        if _valid_github_user(username):
            candidates.append(username)
    if not candidates:
        return ""
    username = max(enumerate(candidates), key=lambda item: (len(item[1]), -item[0]))[1]
    return f"https://github.com/{username}"


def _valid_github_user(username: str) -> bool:
    lowered = username.lower()
    return lowered not in GITHUB_RESERVED and lowered not in PLACEHOLDER_USERS


def _first_linkedin(text: str) -> str:
    candidates: list[tuple[str, str]] = []
    for match in LINKEDIN_RE.finditer(text):
        kind, slug = match.group(1).lower(), match.group(2)
        if slug.lower() in PLACEHOLDER_USERS:
            continue
        if not slug or slug.lower() in {"in", "pub"}:
            continue
        candidates.append((kind, slug.rstrip("_")))
    if not candidates:
        return ""
    kind, slug = max(
        enumerate(candidates),
        key=lambda item: (len(item[1][1]), -item[0]),
    )[1]
    return f"https://www.linkedin.com/{kind}/{slug}"


def _first_name(text: str, filename: str) -> str:
    lines = [_clean_line(raw) for raw in text.splitlines()]
    lines = [line for line in lines if line][:24]
    i = 0
    while i < len(lines):
        line = lines[i]
        if _skip_name_line(line):
            i += 1
            continue
        words = _name_tokens(line)
        if not words:
            i += 1
            continue
        joined = list(words)
        j = i + 1
        while j < len(lines) and len(joined) < 5:
            nxt = lines[j]
            if _skip_name_line(nxt):
                break
            nxt_words = _name_tokens(nxt)
            if len(nxt_words) == 1:
                joined.append(nxt_words[0])
                j += 1
                continue
            break
        if 2 <= len(joined) <= 5:
            candidate = " ".join(joined)
            if candidate.lower() not in PLACEHOLDER_NAMES:
                return _format_name(candidate)
        if len(joined) == 1 and _is_name_word(joined[0]):
            later = _two_word_name(lines[i + 1 :])
            if later:
                return later
            if joined[0].lower() not in NAME_STOPWORDS:
                return _format_name(joined[0])
        i += 1
    return _name_from_filename(filename)


def _two_word_name(lines: list[str]) -> str:
    for line in lines[:8]:
        if _skip_name_line(line):
            continue
        words = _name_tokens(line)
        if 2 <= len(words) <= 5:
            candidate = " ".join(words)
            if candidate.lower() not in PLACEHOLDER_NAMES:
                return _format_name(candidate)
    return ""


def _clean_line(line: str) -> str:
    return line.strip(" |-•\t\ufeff")


def _skip_name_line(line: str) -> bool:
    if not line:
        return True
    if EMAIL_RE.search(line) or PHONE_RE.search(line):
        return True
    if line.lower() in PLACEHOLDER_NAMES:
        return True
    if len(line) > 60:
        return True
    if _name_tokens(line):
        return False
    return bool(SKIP_NAME_RE.search(line))


def _name_tokens(line: str) -> list[str]:
    if "," in line:
        return []
    tokens = [re.sub(r"[^\w'.-]", "", token) for token in line.split()]
    tokens = [token for token in tokens if token]
    if not tokens or len(tokens) > 5:
        return []
    if not all(_is_name_word(token) for token in tokens):
        return []
    return tokens


def _is_name_word(word: str) -> bool:
    cleaned = word.strip(".,")
    if len(cleaned) < 2 or len(cleaned) > 24:
        return False
    if cleaned.lower() in NAME_STOPWORDS or cleaned.lower() in MONTHS:
        return False
    if cleaned.isupper() and len(cleaned) == 2:
        return False
    letters = re.sub(r"[.'’-]", "", cleaned)
    if not letters or not letters.isalpha():
        return False
    return cleaned[:1].isupper() or cleaned.isupper()


def _format_name(name: str) -> str:
    name = re.sub(r"\s+", " ", name).strip()
    if name.isupper() or name.islower():
        return name.title()
    return name


def _name_from_filename(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"(?i)\.docx$", "", stem)
    stem = re.sub(r"(?i)[_-]?resume[_-]?", " ", stem)
    stem = re.sub(r"(?i)\b(cv|curriculum|vitae)\b", " ", stem)
    stem = stem.replace("_", " ").replace("-", " ")
    stem = re.sub(r"\(\s*\d+\s*\)", " ", stem)
    stem = re.sub(
        r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)"
        r"[a-z]*\b",
        " ",
        stem,
        flags=re.I,
    )
    stem = re.sub(r"\b(19|20)\d{2}\b", " ", stem)
    stem = re.sub(r"\b\d{1,2}(st|nd|rd|th)?\b", " ", stem)
    stem = re.sub(r"[^A-Za-z ]", " ", stem)
    words = [
        word
        for word in stem.split()
        if word.lower() not in GENERIC_FILENAME_WORDS and len(word) > 1
    ]
    if not words:
        return ""
    return " ".join(words).title()
