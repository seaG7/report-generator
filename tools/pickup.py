from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from tools.new_work import NewWorkError, create_work
from tools.paths import repo_root


class PickupError(Exception):
    pass


@dataclass
class Pending:
    subject: str
    pdf: Path


_LAB_DIR = re.compile(r"^lab-\d+-")


def slug_from_filename(name: str) -> str:
    stem = Path(name).stem
    slug = re.sub(r"[^\w]+", "-", stem, flags=re.UNICODE).strip("-").lower()
    slug = re.sub(r"-{2,}", "-", slug)
    return (slug or "assignment")[:40]


def list_pending(repo: Path | None = None) -> list[Pending]:
    root = Path(repo) if repo is not None else repo_root()
    subjects = root / "subjects"
    if not subjects.is_dir():
        return []
    pending: list[Pending] = []
    for subj in sorted(subjects.iterdir()):
        if not subj.is_dir() or subj.name.startswith("."):
            continue
        inbox = subj / "inbox"
        if inbox.is_dir():
            for pdf in sorted(inbox.glob("*.pdf")):
                pending.append(Pending(subject=subj.name, pdf=pdf))
        for pdf in sorted(subj.glob("*.pdf")):
            if pdf.parent == subj:
                pending.append(Pending(subject=subj.name, pdf=pdf))
    return pending


def claim_pending(item: Pending, repo: Path | None = None) -> Path:
    root = Path(repo) if repo is not None else repo_root()
    try:
        return create_work(
            item.subject,
            slug_from_filename(item.pdf.name),
            item.pdf,
            move_pdf=True,
            repo=root,
        )
    except NewWorkError as exc:
        raise PickupError(str(exc)) from exc


def claim_all(repo: Path | None = None) -> list[Path]:
    works: list[Path] = []
    for item in list_pending(repo):
        works.append(claim_pending(item, repo=repo))
    return works


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--repo")
    p.add_argument("--json", action="store_true")
    args = p.parse_args()
    repo = Path(args.repo) if args.repo else None
    try:
        works = claim_all(repo)
    except PickupError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    if not works:
        print(
            "Нет PDF. Положите файл задания в subjects/<предмет>/inbox/",
            file=sys.stderr,
        )
        return 2
    if args.json:
        print(json.dumps([str(w) for w in works], ensure_ascii=False))
    else:
        for w in works:
            print(w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
