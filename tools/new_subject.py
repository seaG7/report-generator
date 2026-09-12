from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from tools.paths import repo_root


class SubjectError(Exception):
    pass


_SLUG = re.compile(r"^[a-z0-9]([a-z0-9-]*[a-z0-9])?$")


def new_subject(slug: str, repo: Path | None = None) -> Path:
    slug = slug.strip().lower()
    if not _SLUG.match(slug):
        raise SubjectError(f"Некорректный slug предмета: {slug}")
    root = Path(repo) if repo is not None else repo_root()
    dest = root / "subjects" / slug
    dest.mkdir(parents=True, exist_ok=True)
    inbox = dest / "inbox"
    inbox.mkdir(exist_ok=True)
    gitkeep = inbox / ".gitkeep"
    if not gitkeep.is_file():
        gitkeep.write_text("", encoding="utf-8")
    return dest


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--slug", required=True)
    p.add_argument("--repo")
    args = p.parse_args()
    try:
        dest = new_subject(args.slug, repo=Path(args.repo) if args.repo else None)
    except SubjectError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(dest)
    return 0


if __name__ == "__main__":
    sys.exit(main())
