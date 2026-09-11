from tools.paths import repo_root


def test_three_skills_have_names():
    root = repo_root() / ".grok" / "skills"
    for name in ("report", "pdf-to-md", "build-report"):
        text = (root / name / "SKILL.md").read_text(encoding="utf-8")
        assert text.startswith("---")
        assert f"name: {name}" in text
        assert "description:" in text
