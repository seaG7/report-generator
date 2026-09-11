from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


class DiagramError(Exception):
    pass


def _default_runner(cmd: list[str]) -> int:
    return subprocess.run(cmd, check=False).returncode


def render_mermaid(
    mmd_path: Path,
    png_path: Path | None = None,
    *,
    runner=None,
) -> Path:
    mmd_path = Path(mmd_path)
    png_path = Path(png_path) if png_path else mmd_path.with_suffix(".png")
    run = runner or _default_runner
    cmd = [
        "npx",
        "--yes",
        "@mermaid-js/mermaid-cli",
        "-i",
        str(mmd_path),
        "-o",
        str(png_path),
        "-b",
        "white",
    ]
    if run(cmd) != 0:
        raise DiagramError("Не удалось собрать диаграмму Mermaid")
    return png_path


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("mmd")
    p.add_argument("--out")
    args = p.parse_args()
    try:
        out = render_mermaid(Path(args.mmd), Path(args.out) if args.out else None)
    except DiagramError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
