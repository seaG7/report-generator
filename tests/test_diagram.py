from pathlib import Path
import pytest
from tools.diagram import DiagramError, render_mermaid


def test_render_mermaid_uses_runner_and_writes_png(tmp_path: Path):
    mmd = tmp_path / "a.mmd"
    mmd.write_text("flowchart LR\nA-->B\n", encoding="utf-8")
    png = tmp_path / "a.png"

    def runner(cmd):
        assert cmd[0] == "npx"
        assert str(mmd) in cmd and str(png) in cmd
        png.write_bytes(b"\x89PNG")
        return 0

    out = render_mermaid(mmd, png, runner=runner)
    assert out == png and png.is_file()


def test_render_mermaid_nonzero_raises(tmp_path: Path):
    mmd = tmp_path / "a.mmd"
    mmd.write_text("x", encoding="utf-8")
    with pytest.raises(DiagramError, match="Mermaid"):
        render_mermaid(mmd, runner=lambda cmd: 1)
