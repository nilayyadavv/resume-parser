import os
from pathlib import Path

import pytest

from resume_parser.extract import ParsedResume
from resume_parser.sheets import (
    CredentialsNotFoundError,
    resolve_credentials,
    write_csv,
)


def test_write_csv(tmp_path: Path):
    results = [
        ParsedResume(
            name="Maya Chen",
            email="maya.chen@example.com",
            github="https://github.com/mayachen",
            linkedin="https://www.linkedin.com/in/maya-chen",
            source_file="maya-chen.pdf",
            status="ok",
        )
    ]
    path = write_csv(tmp_path / "out.csv", results)
    content = path.read_text(encoding="utf-8")
    assert "Name,Email,GitHub,LinkedIn,Source file,Status" in content
    assert "Maya Chen" in content
    assert "https://github.com/mayachen" in content


def test_resolve_credentials_missing_explicit_path(tmp_path: Path):
    missing = tmp_path / "nope.json"
    with pytest.raises(CredentialsNotFoundError):
        resolve_credentials(str(missing))


def test_resolve_credentials_env_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    path = tmp_path / "creds.json"
    path.write_text("{}", encoding="utf-8")
    monkeypatch.setenv("GOOGLE_APPLICATION_CREDENTIALS", str(path))
    assert resolve_credentials(None) == path


def test_resolve_credentials_absent_env(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.delenv("GOOGLE_APPLICATION_CREDENTIALS", raising=False)
    assert resolve_credentials(None) is None
