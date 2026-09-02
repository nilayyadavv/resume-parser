from resume_parser.extract import parse_text


def test_extracts_name_email_github_linkedin():
    text = """
    MAYA CHEN
    maya.chen@example.com
    https://github.com/mayachen
    https://www.linkedin.com/in/maya-chen
    Software engineer
    """
    result = parse_text(text, filename="maya-chen.pdf")
    assert result.name == "Maya Chen"
    assert result.email == "maya.chen@example.com"
    assert result.github == "https://github.com/mayachen"
    assert result.linkedin == "https://www.linkedin.com/in/maya-chen"
    assert result.status == "ok"


def test_github_without_scheme_and_linkedin_pub():
    text = """
    Jordan Patel
    jordan.patel@example.com
    github.com/jordanpatel/dotfiles
    linkedin.com/pub/jordan-patel
    """
    result = parse_text(text, filename="jordan.docx")
    assert result.github == "https://github.com/jordanpatel"
    assert result.linkedin == "https://www.linkedin.com/pub/jordan-patel"


def test_skips_reserved_github_paths():
    text = """
    Sam Rivera
    sam.rivera@example.net
    https://github.com/features
    https://github.com/samrivera
    """
    result = parse_text(text, filename="sam.pdf")
    assert result.github == "https://github.com/samrivera"


def test_no_links_status_and_filename_fallback_name():
    text = "Contact me at alex.nguyen@example.org for roles in platform engineering."
    result = parse_text(text, filename="alex-nguyen-resume.docx")
    assert result.email == "alex.nguyen@example.org"
    assert result.github == ""
    assert result.linkedin == ""
    assert result.name == "Alex Nguyen"
    assert result.status == "no GitHub or LinkedIn found"


def test_empty_text():
    result = parse_text("   \n  ", filename="blank.pdf")
    assert result.status == "no extractable text"
